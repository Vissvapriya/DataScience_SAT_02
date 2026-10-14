# DS Sprint Skill-Based Assessment 2 — Use Case 08

**Project:** Bank Customer Insight  
**Organization:** TrustWave Bank | **Sector:** Banking  
**Business question:** How do customers use banking services, and which segments show changing engagement?

> All records and findings are reproducible synthetic data (seed `20261013`). They do not describe real bank customers.

## Setup

From this folder, create and activate a virtual environment, then install the project dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## Run the complete workflow

```powershell
python main.py
```

This creates the five source CSVs in `data/raw/`, cleaned and aggregated tables in `data/processed/`, seven EDA charts, actual holdout model metrics and predictions, an Excel dashboard, assessment/report documents, and a PowerPoint pitch.

The stages can also be run individually from the project root:

```powershell
python -m notebooks.round1_eda
python -m notebooks.round2_machine_learning
python -m notebooks.round3_visualization
```

To regenerate the editable assessment answers and the evidence-checked PDF after running the workflow:

```powershell
python reports/build_assessment_pdf.py
```

## Open the interactive dashboard

Run the complete workflow once, then:

```powershell
streamlit run app.py
```

Filter the dashboard by age band and region to explore engagement, channel use, account activity, complaints, and service trends.

## Project map

- `data/raw/`: five generated source CSVs (`customers`, `accounts`, `transactions`, `channels`, `service`).
- `data/processed/`: standardized source tables, feature table, target labels, and analysis summaries.
- `notebooks/round1_eda.py`: synthetic data, quality inspection, cleaning, and six business-question analyses.
- `notebooks/round2_machine_learning.py`: customer-level features, preprocessing pipelines, and model evaluation.
- `notebooks/round3_visualization.py`: formatted Excel workbook, assessment answers, report, and presentation.
- `outputs/charts/`: channel, age, account, complaint, balance, transaction, and service charts.
- `outputs/metrics/`: before/after quality report, computed EDA answers, classifier metrics, class report, confusion matrix, predictions, and selected model.
- `excel/`, `presentation/`, `reports/`: final assessment deliverables.
- `app.py`: Streamlit prototype.

## Method notes

Transaction duplicates are removed by the generated unique transaction ID, not by amount/date similarity. Amount outliers use IQR fences: records are retained, flagged, and capped for downstream analysis; the original amount remains available. Incomplete service records use explicit missing-category/status labels, and missing resolution time is median-imputed.

The three-class engagement target is simulated from a latent propensity used to generate correlated activity behavior. The latent value is excluded from model features. A stratified 75/25 split and preprocessing pipelines are shared across a majority-class `DummyClassifier`, Logistic Regression, and Random Forest. The selected model maximizes test macro F1, with accuracy as a tie-breaker. Performance is illustrative and must not be interpreted as a real banking model validation.

OpenPyXL does not generate native Excel PivotTables. The workbook dashboard contains instructions and suggested source fields for creating them in desktop Excel using **Insert > PivotTable**.
