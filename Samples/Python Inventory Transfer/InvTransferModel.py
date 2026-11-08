from scipy.optimize._lsq.common import print_header_linear
from connector_blob_tools import read_connector_from_blob, normalize_connection_string
from urllib.parse import quote_plus
import pandas as pd
import sqlalchemy
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.seasonal import STL
import re
from sklearn.metrics import mean_absolute_error, mean_squared_error
from datetime import datetime
import time
import pyodbc

start = time.time()


# Step 1: Connectors
hfsmiscconnector = r"E:\config\HFSMiscProd.bin"
foodproconnector = r"E:\config\ReportingFoodproProd.bin"
hfsmiscdsn = ""
foodprodsn = ""

try:
    hfsmiscdsn = read_connector_from_blob(hfsmiscconnector)
    print("Connector string decrypted:")
except Exception as e:
    print(f"Failed to read blob: {e}")

try:
    foodprodsn = read_connector_from_blob(foodproconnector)
    print("Connector string decrypted:")
except Exception as e:
    print(f"Failed to read blob: {e}")

hfsmiscdsn = (
    "Driver={ODBC Driver 17 for SQL Server};" + hfsmiscdsn
)
hfsmiscdsn = normalize_connection_string(hfsmiscdsn)
foodprodsn = (
    "Driver={ODBC Driver 17 for SQL Server};" + foodprodsn
)
hfsmiscdsn = normalize_connection_string(hfsmiscdsn)
foodprodsn = normalize_connection_string(foodprodsn)

# ########### Get data from HFSMisc ############

encoded_dsn = quote_plus(hfsmiscdsn)

# Step 2: Create the SQLAlchemy engine
engine = sqlalchemy.create_engine(f"mssql+pyodbc:///?odbc_connect={encoded_dsn}")

# Step 3: Query your data into a DataFrame
query = "SELECT * FROM [HFS_Miscellaneous].[dbo].[tab_FOODPRO_TransfersInventory]"
df = pd.read_sql(query, engine)

# Force consistent types
df['ItemNumber'] = df['ItemNumber'].astype(str).str.zfill(4).astype(int)
df['QuantityShipped'] = pd.to_numeric(df['QuantityShipped'], errors='coerce')
df['TransferDate'] = pd.to_datetime(df['TransferDate'].astype(str), format='%Y%m%d', errors='coerce')

# Majority unit shipped
majority_unit_df = (
    df.groupby(['ItemNumber', 'UnitShipped'])
    .size()
    .reset_index(name='Count')
    .sort_values(['ItemNumber', 'Count'], ascending=[True, False])
    .drop_duplicates('ItemNumber')
    .rename(columns={'UnitShipped': 'MajorityUnitShipped'})
)

# Merge in majority unit
df = df.merge(majority_unit_df[['ItemNumber', 'MajorityUnitShipped']], on='ItemNumber', how='left')

# Filter
df_majority = df[df['UnitShipped'] == df['MajorityUnitShipped']].copy()

# Assign week start
df_majority['WeekStartDate'] = df_majority['TransferDate'].apply(
    lambda x: x.to_period('W').start_time if pd.notna(x) else pd.NaT
)

# Deduplicate on key fields to avoid overcounting
df_majority = df_majority.drop_duplicates(subset=[
    'ItemNumber', 'ItemName', 'TransferDate', 'QuantityShipped'
])

# Aggregate by week
weekly_usage = (
    df_majority
    .groupby(['ItemNumber', 'ItemName', 'MajorityUnitShipped', 'WeekStartDate'])
    .agg(TotalQuantityShipped=('QuantityShipped', 'sum'))
    .reset_index()
)

# Full timeline range
full_weeks = pd.date_range(
    start=weekly_usage['WeekStartDate'].min(),
    end=weekly_usage['WeekStartDate'].max(),
    freq='W-MON'
)

# Reconstruct full_weekly_usage for decomposition (not used in forecast loop)
full_weeks = pd.date_range(
    start=weekly_usage['WeekStartDate'].min(),
    end=weekly_usage['WeekStartDate'].max(),
    freq='W-MON'
)

full_ts_list = []

for item_id, group in weekly_usage.groupby('ItemNumber'):
    item_name = group['ItemName'].iloc[0]
    unit = group['MajorityUnitShipped'].iloc[0]

    item_weeks = pd.DataFrame({'WeekStartDate': full_weeks})
    item_weeks['ItemNumber'] = item_id
    item_weeks['ItemName'] = item_name
    item_weeks['MajorityUnitShipped'] = unit

    merged = pd.merge(
        item_weeks, group,
        on=['ItemNumber', 'ItemName', 'MajorityUnitShipped', 'WeekStartDate'],
        how='left'
    )
    merged['TotalQuantityShipped'] = merged['TotalQuantityShipped'].fillna(0)
    full_ts_list.append(merged)

full_weekly_usage = pd.concat(full_ts_list, ignore_index=True)


# Forecast for all items
results = []
now = datetime.now()

for item_id in weekly_usage['ItemNumber'].unique():
    group = weekly_usage[weekly_usage['ItemNumber'] == item_id]
    if group.empty:
        continue

    item_name = group['ItemName'].iloc[0]
    unit = group['MajorityUnitShipped'].iloc[0]

    item_weeks = pd.DataFrame({'WeekStartDate': full_weeks})
    item_weeks['ItemNumber'] = item_id
    item_weeks['ItemName'] = item_name
    item_weeks['MajorityUnitShipped'] = unit

    merged = pd.merge(
        item_weeks, group,
        on=['ItemNumber', 'ItemName', 'MajorityUnitShipped', 'WeekStartDate'],
        how='left'
    )
    merged['TotalQuantityShipped'] = merged['TotalQuantityShipped'].fillna(0)

    ts = merged.set_index('WeekStartDate')['TotalQuantityShipped']

    if ts.sum() > 0 and len(ts) >= 3:
        try:
            model = ExponentialSmoothing(
                ts,
                trend='add',
                seasonal='add',
                seasonal_periods=13,
                initialization_method='estimated'
            ).fit()

            forecast = model.forecast(52)
            forecast_index = pd.date_range(
                start=ts.index[-1] + pd.Timedelta(weeks=1),
                periods=52,
                freq='W-MON'
            )

            hist_df = pd.DataFrame({
                'ItemNumber': item_id,
                'ItemName': item_name,
                'MajorityUnitShipped': unit,
                'WeekStartDate': ts.index,
                'ActualQuantity': ts.values,
                'FittedQuantity': model.fittedvalues
            })
            hist_df['Residual'] = hist_df['ActualQuantity'] - hist_df['FittedQuantity']
            hist_df['4WeekAvg'] = hist_df['ActualQuantity'].rolling(window=4).mean()
            hist_df['ForecastQuantity'] = np.nan
            hist_df['Forecast4WeekAvg'] = np.nan

            forecast_df = pd.DataFrame({
                'ItemNumber': item_id,
                'ItemName': item_name,
                'MajorityUnitShipped': unit,
                'WeekStartDate': forecast_index,
                'ActualQuantity': np.nan,
                'FittedQuantity': np.nan,
                'Residual': np.nan,
                '4WeekAvg': np.nan,
                'ForecastQuantity': forecast.values
            })
            forecast_df['Forecast4WeekAvg'] = forecast_df['ForecastQuantity'].rolling(window=4).mean()

            final_df = pd.concat([hist_df, forecast_df], ignore_index=True)
            final_df['DateAdded'] = now

            results.append(final_df)

        except Exception as e:
            print(f"⚠️ Forecast failed for item {item_id}: {e}")
            continue

# Export everything
if results:
    final = pd.concat(results, ignore_index=True)
    final.to_csv(r"E:\Flatfiles\PythonModels\ABS_TS_INV_FORECASTS.csv", index=False)
    print("✅ Forecast written for all items.")
else:
    print("⚠️ No valid items to forecast.")


from statsmodels.tsa.seasonal import STL

def decompose_all_items(df, seasonal_periods=13):
    results = []

    for item_id in df['ItemNumber'].unique():
        item_df = df[df['ItemNumber'] == item_id].sort_values('WeekStartDate')
        ts = item_df.set_index('WeekStartDate')['TotalQuantityShipped']

        if len(ts) < seasonal_periods * 2 or ts.sum() == 0:
            continue

        try:
            stl = STL(ts, period=seasonal_periods, robust=True).fit()

            decomp_df = pd.DataFrame({
                'ItemNumber': item_id,
                'ItemName': item_df['ItemName'].iloc[0],
                'MajorityUnitShipped': item_df['MajorityUnitShipped'].iloc[0],
                'WeekStartDate': ts.index,
                'Observed': ts.values,
                'Trend': stl.trend.values,
                'Seasonal': stl.seasonal.values,
                'Residual': stl.resid.values,
                'DateAdded': datetime.now()
            })

            results.append(decomp_df)

        except Exception as e:
            print(f"⚠️ Decomposition failed for item {item_id}: {e}")
            continue

    return pd.concat(results, ignore_index=True) if results else pd.DataFrame()

decomp_df = decompose_all_items(full_weekly_usage)
decomp_df.to_csv(r"E:\Flatfiles\PythonModels\ABS_TS_INV_DECOMP.csv", index=False)
print("✅ Decomposition written to CSV.")

from sklearn.metrics import mean_absolute_error, mean_squared_error

def evaluate_holt_winters_models(full_forecast_df):
    summary = []

    for item_id in full_forecast_df['ItemNumber'].dropna().unique():
        item_df = full_forecast_df[full_forecast_df['ItemNumber'] == item_id].copy()

        hist = item_df.dropna(subset=['ActualQuantity', 'FittedQuantity'])

        if hist.empty or hist['ActualQuantity'].sum() == 0:
            continue

        y_true = hist['ActualQuantity'].values
        y_pred = hist['FittedQuantity'].values

        mae = mean_absolute_error(y_true, y_pred)
        mape = np.mean(np.abs((y_true - y_pred) / np.where(y_true == 0, 1, y_true))) * 100
        rmse = mean_squared_error(y_true, y_pred) ** 0.5

        summary.append({
            'ItemNumber': item_id,
            'ItemName': item_df['ItemName'].iloc[0],
            'MajorityUnitShipped': item_df['MajorityUnitShipped'].iloc[0],
            'MAE': mae,
            'MAPE': mape,
            'RMSE': rmse,
            'DateAdded': datetime.now()
        })

    return pd.DataFrame(summary)

error_summary_df = evaluate_holt_winters_models(final)
error_summary_df.to_csv(r"E:\Flatfiles\PythonModels\ABS_TS_INV_MODELERRORS.csv", index=False)
print("✅ Model error summary written to CSV.")
