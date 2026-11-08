# Modeling & Analytics Samples

Brief index of selected Python and R coursework, applied analytics, and statistical modeling projects. The folders contain source files plus companion notebooks, PDFs, and DOCX reports.

## At a glance

| Area | Technology and content | Statistical / machine-learning methods |
|---|---|---|
| [Capstone — SAS Viya and Python](<Samples/Capstone SAS Viya and Python>) | Python, SQL Server, Excel, Word/PDF reporting; retail and convenience-store operations | ARIMAX/SARIMAX, STL decomposition, negative-binomial regression, Apriori-style association rules, descriptive trend analysis |
| [Python — Inventory Transfer](<Samples/Python Inventory Transfer>) | Python, SQL Server/ODBC, pandas, statsmodels | Holt-Winters exponential smoothing, STL decomposition, MAE/RMSE evaluation; weekly item-usage forecasting |
| [Python — Classification](<Samples/Python Courework/Classification>) | Jupyter, pandas, scikit-learn | SVM with linear/RBF/polynomial kernels, Random Forest, AdaBoost |
| [Python — Clustering](<Samples/Python Courework/Clustering>) | Jupyter, pandas, scikit-learn, SciPy | K-means, agglomerative clustering with average/complete/Ward linkage, DBSCAN, clustering metrics |
| [Python — Regression](<Samples/Python Courework/Linear Regression Regression Tree Neural Network>) | Jupyter, pandas, scikit-learn, UCI data | Linear regression, regression tree, MLP neural network, grouped cross-validation |
| [Python — Classification comparison](<Samples/Python Courework/Logistic Regression Bayes Neural Networks>) | Jupyter, pandas, scikit-learn | Logistic regression, Gaussian Naive Bayes, decision tree, MLP neural network |
| [Python — Ensemble-model project](<Samples/Python Courework/Final Project Ensemble Models>) | Jupyter, pandas, scikit-learn, seaborn | Customer ad-click propensity modeling with logistic regression, decision tree, Naive Bayes, MLP, and Random Forest |
| [Python — Text analysis](<Samples/Python Courework/Text Analysis>) | Python, NLTK, scikit-learn | Text normalization, tokenization, stemming, CountVectorizer, Naive Bayes, linear SVM |
| [Python — Association rules](<Samples/Python Courework/Association Rules>) | Python, pandas, scikit-learn, NetworkX, Matplotlib | TF-IDF/cosine and Jaccard similarity; Apriori itemsets, support, confidence, lift, and rule-network visualization |
| [Python — GSS text/categorical modeling](<Samples/Python Text Classification>) | Jupyter, pandas, Pyreadstat, SciPy, statsmodels | Exploratory analysis, chi-square tests, logistic regression, odds ratios, confidence intervals |
| [Python — Fabric reporting notebook](<Samples/Fabric Python Notebook Experiment ID Printing>) | Jupyter, pandas, NumPy, SciPy, Matplotlib | Operational reporting, daily/weekly aggregation, moving averages, descriptive statistics, KDE/distribution plots |
| [R — Linear regression](<Samples/R Coursework/Linear Regression>) | R Markdown, base R, dplyr, ggplot2 | `lm()` models, residual diagnostics, Q-Q plots, scale-location, leverage/Cook's distance |
| [R — Classification and neural networks](<Samples/R Coursework/Neural Networks C5.0 CART>) | R Markdown, caret, C50, rpart, neuralnet/nnet, pROC | C5.0, CART, feed-forward neural networks, ROC/AUC, confusion matrices, variable importance |
| [R — Bayes/logistic regression](<Samples/R Coursework/Bayes Logistic Regression>) | R Markdown, e1071, caret, pROC, tidyverse | Naive Bayes, logistic regression comparison, ROC/AUC, posterior/likelihood illustrations |
| [R — Time series](<Samples/R Coursework/Time Series>) | R Markdown, forecast, tseries, ggplot2 | Moving-average and exponential-smoothing forecasts, seasonal decomposition, ACF/PACF diagnostics |

## File index

### Python projects

#### Capstone — SAS Viya and Python

- `Arimax_cstore_detailed_combined.py` and `Arimax_cstore_detailed_combined_date_adjusted.py` — weekly C-store analysis by location and all-store total; AR(1)-style ARIMAX/SARIMAX models using a legacy-versus-Mashgin indicator, STL decomposition, residual diagnostics, charts, and report export.
- `Mashgin_15min_overall_analysis.py` — SQL Server extraction and 15-minute transaction-flow analysis; negative-binomial regressions with store/time controls and Word/Excel/PNG outputs.
- `mashgin_bundle_apriori_full.py` — item and category basket analysis using support, confidence, lift, and incremental-basket metrics.
- `mashgin_strategic_visualizations.py` — executive visualizations for the association-rule outputs.
- `01Capstone_Part_I.pdf`, `02Capstone_Part_II.pdf`, and `M3_Mashgin_Final_Recommendations.docx` — capstone documentation and recommendations.

#### Coursework and applied projects

- [Inventory Transfer](<Samples/Python Inventory Transfer/InvTransferModel.py>) — extracts inventory-transfer data, aggregates weekly usage, forecasts 52 weeks per item with additive Holt-Winters smoothing, and evaluates/decomposes the resulting series. `Item Usage Predictions.pdf` is the companion report.
- [Classification](<Samples/Python Courework/Classification/L10_DAAN862.ipynb>) — breast-cancer classification comparison across SVM kernels, Random Forest, and AdaBoost.
- [Clustering](<Samples/Python Courework/Clustering/L11_DAAN862.ipynb>) — wheat-seed clustering with K-means, hierarchical methods, and tuned DBSCAN.
- [Regression](<Samples/Python Courework/Linear Regression Regression Tree Neural Network/L9_DAAN862%20.ipynb>) — Parkinson’s total-UPDRS prediction using linear regression, a regression tree, and an MLP regressor with grouped cross-validation.
- [Logistic/Bayes/neural networks](<Samples/Python Courework/Logistic Regression Bayes Neural Networks/L8_DAAN862.ipynb>) — breast-cancer classification using logistic regression, Gaussian Naive Bayes, decision tree, and MLP models.
- [Ensemble-model final project](<Samples/Python Courework/Final Project Ensemble Models/DAAN862_E_Smith_Final.ipynb>) — customer email-ad click propensity; compares and tunes several classification models and reports accuracy/AUC-style metrics.
- [Text analysis](<Samples/Python Courework/Text Analysis/L12_DAAN862.ipynb>) and [lesson 13 script](<Samples/Python Courework/Text Analysis/lesson%2013.py>) — regular expressions, word counts, NLTK preprocessing, Naive Bayes, and linear SVM text classification.
- [Association rules](<Samples/Python Courework/Association Rules/Pythonanlaysis.py>) — course/student data cleaning plus similarity analysis; the related scripts generate Apriori rule tables and NetworkX visualizations. `E_Smith_term_project.pdf` is the companion report.
- [GSS categorical modeling](<Samples/Python Text Classification/GSSCategoricalModeling.ipynb>) — GSS 2024 science-skepticism analysis with chi-square screening, logistic regression, odds ratios, and forest plots.
- [Fabric notebook](<Samples/Fabric Python Notebook Experiment ID Printing/ABSIT_Process_Reporting_291_DE_NB_IDMS_Printing%20(1).ipynb>) — print-volume reporting by date, region, and location with moving-average and distribution analysis. `Notebook.pdf` is the rendered companion.

### R projects

- [Linear Regression](<Samples/R Coursework/Linear Regression/HW5_IE575_001_E_SMITH.Rmd>) — Anscombe-data regression comparison and diagnostic plots; rendered report: `HW5_IE575_001_E_SMITH.docx`.
- [C5.0 and CART](<Samples/R Coursework/Neural Networks C5.0 CART/HW6_C50.Rmd>) — heart-attack diagnosis classification with full-data and partitioned C5.0 models, plus CART comparison.
- [Neural networks](<Samples/R Coursework/Neural Networks C5.0 CART/HW7.Rmd>) — patient-risk classification with a one-hidden-layer neural network and C5.0 comparison, including ROC/AUC and feature-importance analysis.
- [Credit-default neural network](<Samples/R Coursework/Neural Networks C5.0 CART/HW8.Rmd>) — partitioned `nnet` model for credit-card default, ROC/AUC evaluation, and Garson/correlation-based importance.
- [Bayes/logistic regression](<Samples/R Coursework/Bayes Logistic Regression/HW9.Rmd>) — Naive Bayes classification and logistic-regression comparison with preprocessing, ROC/AUC, and likelihood/posterior illustrations.
- [Time series](<Samples/R Coursework/Time Series/HW10.Rmd>) — Citi Bike demand exploration, daily seasonality/decomposition, ACF/PACF checks, exponential smoothing, and 3-hour/3-day moving-average forecasts.

## Notes

- The PDFs and DOCX files are supporting reports or rendered outputs; the notebooks and scripts are the primary source artifacts.
- Several coursework files reference local datasets or historical working directories that are not included here, so they should be treated as portfolio samples rather than turnkey packages.
- Filenames and folder names are preserved as found in the sample collection.
