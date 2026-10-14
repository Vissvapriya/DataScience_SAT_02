"""Build assessment answers and an evidence-checked PDF from project outputs."""

from __future__ import annotations

import json
import re
import subprocess
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
PDF_PATH = REPORTS / "CH08_Bank_Customer_Insight_Assessment_Answers.pdf"
MARKDOWN_PATH = REPORTS / "assessment_sheet_answers.md"
NAVY = colors.HexColor("#163A5F")
TEAL = colors.HexColor("#168C8C")
PALE = colors.HexColor("#EAF1F7")
INK = colors.HexColor("#202B36")
GREEN = colors.HexColor("#E4F2E7")
AMBER = colors.HexColor("#FFF2D7")
ROSE = colors.HexColor("#FBE7E7")

EXPECTED_DATASETS = [
    "customers",
    "accounts",
    "transactions",
    "channels",
    "service",
]
CHART_TITLES = {
    "channel_usage.png": "Banking-channel usage",
    "age_group_comparison.png": "Average channel usage by age band",
    "account_activity.png": "Account activity, Q1 vs Q4",
    "regional_complaints.png": "Complaint counts by region",
    "balance_distribution.png": "Balance by engagement segment",
    "transaction_trends.png": "Monthly transaction trend",
    "service_requests.png": "Service requests by type",
}
FEATURE_LABELS = {
    "age_band": "Age band",
    "region": "Region",
    "relationship_months": "Relationship length (months)",
    "account_balance": "Total account balance",
    "account_count": "Account count",
    "transaction_count": "Transaction count",
    "transaction_frequency_per_month": "Transactions per month",
    "branch_uses": "Branch uses",
    "atm_uses": "ATM uses",
    "mobile_uses": "Mobile uses",
    "internet_uses": "Internet uses",
    "complaint_count": "Complaint count",
    "service_request_count": "Service-request count",
}


def load_evidence() -> dict[str, object]:
    """Load project outputs and verify dataset/workbook facts from disk."""
    findings_path = ROOT / "outputs" / "metrics" / "round1_findings.json"
    metrics_path = ROOT / "outputs" / "metrics" / "model_metrics.json"
    if not findings_path.exists() or not metrics_path.exists():
        raise FileNotFoundError(
            "Run `python main.py` from the project directory before building the assessment PDF."
        )
    findings = json.loads(findings_path.read_text(encoding="utf-8"))
    metrics = json.loads(metrics_path.read_text(encoding="utf-8"))
    datasets = {}
    for name in EXPECTED_DATASETS:
        csv_path = ROOT / "data" / "raw" / f"{name}.csv"
        if not csv_path.exists():
            raise FileNotFoundError(f"Required source dataset is missing: {csv_path}")
        frame = pd.read_csv(csv_path)
        datasets[name] = {
            "records": int(len(frame)),
            "attributes": int(frame.shape[1]),
            "columns": list(frame.columns),
        }
    features = pd.read_csv(ROOT / "data" / "processed" / "customer_insight_features.csv")
    target = pd.read_csv(ROOT / "data" / "processed" / "customer_engagement_target.csv")

    workbook_path = ROOT / "excel" / "bank_customer_insight_dashboard.xlsx"
    if not workbook_path.exists():
        raise FileNotFoundError(f"Required Excel workbook is missing: {workbook_path}")
    workbook = load_workbook(workbook_path, data_only=False)
    formula_count = sum(
        1
        for sheet in workbook.worksheets
        for row in sheet.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    )
    table_count = sum(len(sheet.tables) for sheet in workbook.worksheets)
    chart_count = sum(len(sheet._charts) for sheet in workbook.worksheets)
    pivot_count = sum(len(sheet._pivots) for sheet in workbook.worksheets)
    dashboard = workbook["Summary Dashboard"]
    formulas = [
        cell.value
        for row in dashboard.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    ]
    formula_text = " ".join(formulas).upper()
    has_statistical_formula = any(
        function in formula_text
        for function in ("AVERAGE(", "MEDIAN(", "STDEV(", "COUNTIF(", "SUMIF(")
    )
    repository_url = ""
    try:
        result = subprocess.run(
            ["git", "-C", str(ROOT), "remote", "get-url", "origin"],
            capture_output=True,
            text=True,
            check=False,
            timeout=5,
        )
        if result.returncode == 0:
            repository_url = result.stdout.strip()
    except (OSError, subprocess.TimeoutExpired):
        repository_url = ""
    chart_files = sorted(
        path.name for path in (ROOT / "outputs" / "charts").glob("*.png")
    )
    return {
        "findings": findings,
        "metrics": metrics,
        "datasets": datasets,
        "feature_columns": list(features.columns),
        "feature_rows": len(features),
        "feature_attributes": features.shape[1],
        "target_rows": len(target),
        "target_attributes": target.shape[1],
        "workbook_sheets": workbook.sheetnames,
        "workbook_tables": table_count,
        "workbook_charts": chart_count,
        "workbook_formulas": formula_count,
        "workbook_statistical_formula": has_statistical_formula,
        "workbook_pivots": pivot_count,
        "chart_files": chart_files,
        "repository_url": repository_url,
    }


def _status(evidence: dict[str, object]) -> list[tuple[str, str, str]]:
    charts = evidence["workbook_charts"]
    formulas = evidence["workbook_formulas"]
    pivot_count = evidence["workbook_pivots"]
    return [
        ("Cleaned data imported into Excel", "Complete", "Cleaned customer, account, transaction and service-analysis data are included in workbook sheets."),
        ("Data organization and formatting", "Complete", f"{len(evidence['workbook_sheets'])} sheets; {evidence['workbook_tables']} formatted tables; styled headers and frozen panes."),
        ("Formulas and statistical functions", "Partial", f"{formulas} Excel formulas (COUNTA); no Excel statistical formulas. Several summaries are precomputed in Python."),
        ("Channel usage analysis", "Complete", "Channel Analysis sheet has totals for all four channels."),
        ("Customer segment comparison", "Complete", "Customer Segments and Balance Analysis sheets include engagement classes and segment balances."),
        ("Account activity analysis", "Complete", "Account Activity sheet compares Q1 and Q4 transactions by account type."),
        ("Complaint analysis", "Complete", "Complaint Analysis sheet reports complaint counts by region."),
        ("Balance distribution analysis", "Complete", "Balance Analysis sheet reports balance statistics by segment."),
        ("Service-request trend analysis", "Complete", "Service Trends sheet contains request and complaint counts by request type."),
        ("Charts", "Partial", f"{charts} Excel charts are embedded (channel usage and account activity); other analyses are tables only."),
        ("PivotTables and PivotCharts", "Complete" if pivot_count else "Missing", f"Workbook contains {pivot_count} native PivotTables/PivotCharts; it includes instructions for creating PivotTables in Excel."),
        ("Summary sheet with findings", "Complete", "Summary Dashboard presents findings, formulas, model results, charts, and PivotTable instructions."),
    ]


def _pitch(findings: dict[str, object]) -> str:
    term = findings["account_activity_q1_q4"]["Term Deposit"]
    repeat = findings["repeated_requests_by_service_type"]
    decline_pct = abs(term["change_percent"])
    return (
        "TrustWave Bank needs to understand how customers use its channels and where service "
        "friction may be reducing engagement. In this synthetic sample, mobile banking led "
        f"with {findings['top_channel_uses']:,} reported uses, while North recorded the largest complaint "
        f"count at {findings['regional_complaint_counts']['North']:,}. "
        "These patterns make digital access and regional service quality useful areas for "
        "management review, but counts should be validated against real customer volumes. "
        f"Term Deposit activity was {decline_pct:.2f}% lower in Q4 than Q1, a reason to investigate "
        "whether product journeys or seasonal behavior merit follow-up. Repeat requests "
        f"for Card support and Payments were tied at {repeat['Card support']} customer-service-type pairs each. "
        "We recommend piloting assisted digital onboarding for older customers, reviewing "
        "North complaint causes and repeat-request handling, and testing a Term Deposit "
        "retention contact. Measure adoption, complaint resolution, and repeat contacts "
        "before scaling any intervention."
    )


def build_markdown(evidence: dict[str, object]) -> str:
    findings = evidence["findings"]
    metrics = evidence["metrics"]
    selected_name = metrics["selected_model"]
    selected = metrics["models"][selected_name]
    datasets = evidence["datasets"]
    feature_names = metrics["feature_names"]
    quality_before = findings["data_quality"]["before"]
    quality_after = findings["data_quality"]["after"]
    dataset_rows = "\n".join(
        f"| `{name}.csv` | Synthetic data generated for this project | "
        f"{datasets[name]['records']:,} | {datasets[name]['attributes']} | "
        f"{', '.join(datasets[name]['columns'])} |"
        for name in EXPECTED_DATASETS
    )
    feature_list = ", ".join(
        FEATURE_LABELS.get(name, name.replace("_", " "))
        for name in feature_names
    )
    class_rows = "\n".join(
        f"| {segment} | {selected['classification_report'][segment]['precision']:.4f} | "
        f"{selected['classification_report'][segment]['recall']:.4f} | "
        f"{selected['classification_report'][segment]['f1-score']:.4f} | "
        f"{int(selected['classification_report'][segment]['support'])} |"
        for segment in metrics["class_order"]
    )
    chart_lines = "\n".join(
        f"- **{CHART_TITLES.get(name, name)}** (`outputs/charts/{name}`)."
        for name in evidence["chart_files"]
    )
    checklist = "\n".join(
        f"| {item} | **{status}** | {detail} |"
        for item, status, detail in _status(evidence)
    )
    model_rows = "\n".join(
        f"| {name} | {score['accuracy']:.4f} | {score['macro_f1']:.4f} |"
        for name, score in metrics["models"].items()
    )
    formula_state = (
        "Excel formulas include `COUNTA` counts; no statistical formulas are present."
        if not evidence["workbook_statistical_formula"]
        else "Excel formulas include count and statistical functions."
    )
    raw_records = sum(data["records"] for data in datasets.values())
    project_target_shape = (
        f"{evidence['feature_rows']:,} rows x {evidence['feature_attributes']} feature-table "
        f"columns, plus the separate target ({evidence['target_attributes']} columns)."
    )
    markdown = f"""# DS Sprint Skill-Based Assessment 2 — Use Case 08

**Organization:** TrustWave Bank | **Sector:** Banking  
**Business question:** How do customers use banking services, and which segments show changing engagement?  
**Data note:** All datasets and findings are synthetic (random seed `{findings['seed']}`); no real bank-customer data was used.

## Student Details

- **Name:** Vissvapriya P
- **Register Number:** ____________________________________
- **Class:** III CSE D
- **Challenge ID:** CH08
- **Challenge Title:** Bank Customer Insight

## Round 1 — Data Understanding, Preparation & EDA

### 1. Problem Statement

This project analyzes TrustWave Bank's synthetic customer, account, transaction, channel, and service data to understand engagement, channel preferences, account activity, and complaints. The aim is to identify measurable patterns and recommend ways to improve customer experience and service delivery.

### 2. Dataset Details

Source for each separate source table: synthetic data generated by this project with a fixed seed. Source record and attribute counts below were verified from the CSV files in `data/raw/`.

| Dataset | Source | Records | Attributes | Attributes / fields |
|---|---|---:|---:|---|
{dataset_rows}

The five source CSVs are separate related tables joined using customer and account identifiers; total source-table rows: **{raw_records:,}** (rows across tables, not distinct customers). The generated customer-level analytical feature table has **{project_target_shape}** The model input table joins these features with the separate three-class target.

### 3. Data Preprocessing

- Filled **{findings['missing_channel_values_filled']:,}** missing channel-use cells with the median for the customer's age band; cleaned channel data has no missing cells.
- Removed **{findings['duplicate_transaction_rows_removed']:,}** repeated transaction IDs (**{datasets['transactions']['records']:,}** raw rows to **{quality_after['transactions']['records']:,}** cleaned rows). De-duplication uses transaction identifiers so separately identified similar transactions are retained.
- Standardized inconsistent account labels to Checking, Savings, Credit Card, or Term Deposit.
- Applied IQR outlier fences to transaction amount: lower **${findings['transaction_amount_iqr_bounds']['lower']:.2f}**, upper **${findings['transaction_amount_iqr_bounds']['upper']:.2f}**. Flagged and capped **{findings['transaction_amount_outliers_flagged_and_capped']:,}** amounts; retained each original amount in `amount_original`.
- Filled **{quality_before['service']['missing_cells']:,}** incomplete service cells: missing request types as `Unspecified`, resolution statuses as `Unknown`, and resolution days with the observed median. Customer/account IDs were used for validated joins; transaction count and other features were aggregated to customer level.

| Source table | Before rows | After rows | Before missing cells | After missing cells | Before duplicate rows | After duplicate rows |
|---|---:|---:|---:|---:|---:|---:|
"""
    for name in EXPECTED_DATASETS:
        markdown += (
            f"| {name} | {quality_before[name]['records']:,} | {quality_after[name]['records']:,} | "
            f"{quality_before[name]['missing_cells']:,} | {quality_after[name]['missing_cells']:,} | "
            f"{quality_before[name]['duplicate_rows']:,} | {quality_after[name]['duplicate_rows']:,} |\n"
        )
    markdown += f"""
### 4. Two EDA Findings

1. **Channel preference:** Mobile banking was the most-used channel, with **{findings['top_channel_uses']:,}** reported uses.
2. **Complaints by region:** North had the largest complaint count, with **{findings['regional_complaint_counts']['North']:,}** complaint records.

## Round 2 — Feature Engineering & Machine Learning

### 1. Selected Features

The implemented model uses: **{feature_list}**. These describe demographics and relationship history, products and balances, transactions, channel behavior, and service experience. The target is **`engagement_segment`** with **Low, Medium, and High** classes, generated as tertiles of a simulated latent engagement propensity. The latent propensity itself is excluded from the features.

### 2. ML Algorithm Used

Compared a most-frequent `DummyClassifier`, Logistic Regression, and Random Forest Classifier. **Logistic Regression** was selected by highest held-out macro F1-score (accuracy breaks ties).

### 3. Justification

Logistic Regression supports the three-class classification task and is a useful, comparatively interpretable model for this student prototype. Numeric predictors were standardized and categorical predictors one-hot encoded inside a Scikit-learn pipeline. The comparison used the same stratified 75/25 train/test split (random state 42) for all models.

### 4. Model Performance with actual accuracy and F1-score

**Selected model:** {selected_name} | **Test records:** {metrics['test_records']:,} of {metrics['training_records'] + metrics['test_records']:,}  
**Test accuracy:** {selected['accuracy']:.4f} ({selected['accuracy']:.2%}) | **Test macro F1-score:** {selected['macro_f1']:.4f}

| Compared model | Test accuracy | Test macro F1 |
|---|---:|---:|
{model_rows}

Selected-model class metrics:

| Actual class | Precision | Recall | F1-score | Test records |
|---|---:|---:|---:|---:|
{class_rows}

On this synthetic holdout, the selected model correctly classified **{selected['accuracy']:.2%}** of records and achieved macro F1 **{selected['macro_f1']:.4f}**, with class-level performance varying. This is a demonstration result, not evidence of validated real-world banking performance.

## Round 3 — Intelligence Layer, Visualization & Prototype

### 1. Visualization

The following seven image files exist in `outputs/charts/`:
{chart_lines}

The Excel workbook **`excel/bank_customer_insight_dashboard.xlsx`** also exists, with a Summary Dashboard sheet, nine analysis/data sheets, and embedded charts. The Streamlit app **`app.py`** is implemented with age-band and region filters and displays engagement, channel, account, complaint, and service summaries.

### 2. Key Insights

1. Mobile was the top-use channel (**{findings['top_channel_uses']:,}** uses); the North region recorded the most complaints (**{findings['regional_complaint_counts']['North']:,}**).
2. Term Deposit activity fell from **{findings['account_activity_q1_q4']['Term Deposit']['q1_transactions']:,}** transactions in Q1 to **{findings['account_activity_q1_q4']['Term Deposit']['q4_transactions']:,}** in Q4 2025 (**{findings['account_activity_q1_q4']['Term Deposit']['change_percent']:.2f}%**). Card support and Payments are tied for most repeated request categories, with **{findings['repeated_requests_by_service_type']['Card support']}** customer/service-type pairs each (more than one request).

### 3. Prototype / Intelligence Layer

Python and Pandas clean and join the synthetic tables, create customer-level behavioral features, and evaluate three Scikit-learn classification pipelines. The implemented Streamlit prototype presents filtered customer engagement, channel use, account activity, complaint counts, and service trends. A separate formatted Excel dashboard workbook has also been generated; the PowerPoint pitch deck is in `presentation/`.

### 4. Recommendations

1. Pilot assisted digital onboarding for customers aged **60+**, whose average mobile and internet uses were **{findings['channel_average_uses_by_age_band']['60+']['mobile_uses']:.2f}** and **{findings['channel_average_uses_by_age_band']['60+']['internet_uses']:.2f}** respectively, the lowest combined average among the age bands. Measure adoption and repeat use before scaling.
2. Review complaint causes and resolution processes in **North** (**{findings['regional_complaint_counts']['North']:,}** complaint records); investigate repeat contacts for **Card support and Payments** (**{findings['repeated_requests_by_service_type']['Card support']}** customer/service-type pairs each). Test a targeted Term Deposit follow-up in light of the measured Q1-to-Q4 activity decline.

## Round 3 — Data Analysis Using Excel: Evidence Checklist

Workbook inspected: `excel/bank_customer_insight_dashboard.xlsx`. Completion refers to the actual Excel workbook, not Python outputs alone.

| Assessment requirement | Status | Evidence from workbook |
|---|---|---|
{checklist}

**Pivot note:** OpenPyXL does not create native PivotTables/PivotCharts here; the workbook dashboard contains instructions for creating PivotTables in desktop Excel. Those native pivot deliverables still need to be created if required.

## Round 4 — Pitching Business Insights

{_pitch(findings)}

## Final Submission

- **GitHub Repository Link:** {evidence['repository_url']}
- **Overall Result / Conclusion:** The completed synthetic-data workflow links customer, channel, account, transaction, and service patterns and compares three engagement classifiers. Mobile leads channel use, North has the highest complaint count, and Logistic Regression achieved **{selected['accuracy']:.2%}** test accuracy with macro F1 **{selected['macro_f1']:.4f}**. Validate insights and interventions against real, governed data before operational use.
"""
    return markdown


def _styles() -> dict[str, ParagraphStyle]:
    base = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "AssessmentTitle",
            parent=base["Title"],
            fontName="Helvetica-Bold",
            fontSize=19,
            leading=22,
            textColor=NAVY,
            alignment=TA_CENTER,
            spaceAfter=5,
        ),
        "subtitle": ParagraphStyle(
            "AssessmentSubtitle",
            parent=base["Normal"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#536779"),
            alignment=TA_CENTER,
            spaceAfter=7,
        ),
        "h1": ParagraphStyle(
            "SectionHeading",
            parent=base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=16,
            textColor=NAVY,
            spaceBefore=4,
            spaceAfter=5,
            keepWithNext=True,
        ),
        "h2": ParagraphStyle(
            "QuestionHeading",
            parent=base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=9.5,
            leading=12,
            textColor=TEAL,
            spaceBefore=4,
            spaceAfter=2.5,
            keepWithNext=True,
        ),
        "body": ParagraphStyle(
            "BodyTextCompact",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=8.3,
            leading=11,
            textColor=INK,
            spaceAfter=3.5,
        ),
        "small": ParagraphStyle(
            "SmallText",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=7.2,
            leading=9.1,
            textColor=INK,
            spaceAfter=1,
        ),
        "table": ParagraphStyle(
            "TableText",
            parent=base["BodyText"],
            fontName="Helvetica",
            fontSize=6.9,
            leading=8.5,
            textColor=INK,
        ),
        "table_white": ParagraphStyle(
            "TableHeader",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7,
            leading=8.5,
            textColor=colors.white,
        ),
        "callout": ParagraphStyle(
            "Callout",
            parent=base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.1,
            leading=10.5,
            textColor=NAVY,
        ),
    }


def _p(text: str, style: ParagraphStyle, markup: bool = False) -> Paragraph:
    if markup:
        return Paragraph(text, style)
    safe_text = escape(text)
    safe_text = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", safe_text)
    safe_text = re.sub(r"`([^`]+)`", r'<font name="Courier">\1</font>', safe_text)
    return Paragraph(safe_text, style)


def _table(
    rows: list[list[str]],
    widths: list[float],
    styles: dict[str, ParagraphStyle],
    font_size: float = 7,
    repeat: int = 1,
) -> Table:
    converted = []
    for row_index, row in enumerate(rows):
        para_style = styles["table_white"] if row_index == 0 and repeat else styles["table"]
        converted.append([_p(str(value), para_style) for value in row])
    table = Table(converted, colWidths=widths, repeatRows=repeat, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), NAVY),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#C8D3DD")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, PALE]),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
            ]
        )
    )
    return table


def _footer(canvas, document) -> None:
    canvas.saveState()
    width, height = letter
    canvas.setStrokeColor(colors.HexColor("#D7E0E8"))
    canvas.setLineWidth(0.5)
    canvas.line(document.leftMargin, 0.48 * inch, width - document.rightMargin, 0.48 * inch)
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.HexColor("#617384"))
    canvas.drawString(document.leftMargin, 0.31 * inch, "CH08 | Bank Customer Insight | Synthetic-data assessment")
    canvas.drawRightString(width - document.rightMargin, 0.31 * inch, f"Page {document.page}")
    canvas.restoreState()


def build_pdf(evidence: dict[str, object]) -> None:
    findings = evidence["findings"]
    metrics = evidence["metrics"]
    selected = metrics["models"][metrics["selected_model"]]
    styles = _styles()
    width, _ = letter
    usable_width = width - 1.0 * inch
    story = []

    def heading(level: str, text: str) -> None:
        story.append(_p(text, styles[level]))

    def body(text: str, style: str = "body") -> None:
        story.append(_p(text, styles[style]))

    # Page 1: Student details and Round 1.
    story.append(_p("BANK CUSTOMER INSIGHT", styles["title"]))
    story.append(_p("DS Sprint Skill-Based Assessment 2 | Use Case 08 | TrustWave Bank", styles["subtitle"]))
    body(
        f"Synthetic data generated with fixed seed {findings['seed']}. Findings and model metrics are illustrative and do not describe real bank customers.",
        "small",
    )
    heading("h1", "Student Details")
    student_rows = [
        ["Name", "Vissvapriya P", "Register Number", "________________________________"],
        ["Class", "III CSE D", "Challenge ID", "CH08"],
        ["Challenge Title", "Bank Customer Insight", "Sector", "Banking"],
    ]
    student_table = Table(
        [[_p(cell, styles["table"]) for cell in row] for row in student_rows],
        colWidths=[0.95 * inch, 1.65 * inch, 1.15 * inch, 2.75 * inch],
    )
    student_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), PALE),
                ("BACKGROUND", (2, 0), (2, -1), PALE),
                ("TEXTCOLOR", (0, 0), (0, -1), NAVY),
                ("TEXTCOLOR", (2, 0), (2, -1), NAVY),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C8D3DD")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    story.extend([student_table, Spacer(1, 4)])
    heading("h1", "Round 1 - Data Understanding, Preparation & EDA")
    heading("h2", "1. Problem Statement")
    body(
        "This project analyzes TrustWave Bank's synthetic customer, account, transaction, channel, and service data to understand engagement, channel preferences, account activity, and complaints. The aim is to identify measurable patterns and recommend practical service improvements."
    )
    heading("h2", "2. Dataset Details")
    body("All five source tables are separate linked CSV datasets; counts below are verified directly from the raw CSV files.", "small")
    dataset_rows = [["Dataset", "Records", "Attributes", "Actual fields"]]
    for name in EXPECTED_DATASETS:
        dataset = evidence["datasets"][name]
        dataset_rows.append(
            [
                f"{name}.csv",
                f"{dataset['records']:,}",
                str(dataset["attributes"]),
                ", ".join(dataset["columns"]),
            ]
        )
    story.append(
        _table(
            dataset_rows,
            [0.86 * inch, 0.52 * inch, 0.57 * inch, usable_width - 1.95 * inch],
            styles,
            repeat=1,
        )
    )
    body(
        f"These separate tables total {sum(d['records'] for d in evidence['datasets'].values()):,} source rows (not distinct customers). The merged customer feature table is {evidence['feature_rows']:,} rows by {evidence['feature_attributes']} columns; the target is a separate table.",
        "small",
    )
    heading("h2", "3. Data Preprocessing")
    body(
        f"Filled {findings['missing_channel_values_filled']:,} missing channel values using age-band medians. Standardized account labels. Removed {findings['duplicate_transaction_rows_removed']:,} duplicate transaction IDs ({evidence['datasets']['transactions']['records']:,} raw rows to {findings['data_quality']['after']['transactions']['records']:,} cleaned rows), retaining other uniquely identified similar transactions."
    )
    body(
        f"Used IQR amount fences of ${findings['transaction_amount_iqr_bounds']['lower']:.2f} and ${findings['transaction_amount_iqr_bounds']['upper']:.2f}; flagged and capped {findings['transaction_amount_outliers_flagged_and_capped']:,} values while retaining the original amount. Filled {findings['data_quality']['before']['service']['missing_cells']:,} missing service cells with explicit labels or the observed median. Joined tables on customer/account IDs and aggregated behavioral counts to customer level."
    )
    heading("h2", "4. Two EDA Findings")
    body(
        f"1. **Mobile banking was the most-used channel**, with {findings['top_channel_uses']:,} reported uses."
    )
    body(
        f"2. **North had the highest complaint count**, with {findings['regional_complaint_counts']['North']:,} complaint records."
    )

    # Page 2: Round 2.
    story.append(PageBreak())
    heading("h1", "Round 2 - Feature Engineering & Machine Learning")
    heading("h2", "1. Selected Features")
    feature_text = ", ".join(
        FEATURE_LABELS.get(name, name.replace("_", " "))
        for name in metrics["feature_names"]
    )
    body(f"Implemented inputs: {feature_text}.")
    body(
        "Target: customer engagement segment (Low, Medium, High), defined by tertiles of a simulated latent engagement propensity. The latent value itself is excluded from model input to avoid direct target leakage."
    )
    heading("h2", "2. ML Algorithm Used")
    body(
        f"Compared DummyClassifier baseline, Logistic Regression, and Random Forest Classifier. The selected model is **{metrics['selected_model']}**, based on the highest test macro F1-score (accuracy breaks ties)."
    )
    heading("h2", "3. Justification")
    body(
        "Logistic Regression is a suitable, comparatively interpretable baseline for three-class classification. Standardization for numeric inputs and one-hot encoding for categorical inputs are fitted inside Scikit-learn pipelines, using a shared stratified 75/25 split for a fair comparison."
    )
    heading("h2", "4. Model Performance")
    body(
        f"Holdout test set: {metrics['test_records']:,} customers; training set: {metrics['training_records']:,}. Selected model test accuracy: {selected['accuracy']:.4f} ({selected['accuracy']:.2%}); macro F1-score: {selected['macro_f1']:.4f}."
    )
    model_rows = [["Model", "Test accuracy", "Test macro F1"]]
    model_rows.extend(
        [name, f"{value['accuracy']:.4f}", f"{value['macro_f1']:.4f}"]
        for name, value in metrics["models"].items()
    )
    story.append(_table(model_rows, [2.8 * inch, 1.35 * inch, 1.35 * inch], styles))
    story.append(Spacer(1, 5))
    class_rows = [["Class", "Precision", "Recall", "F1-score", "Support"]]
    class_rows.extend(
        [
            label,
            f"{selected['classification_report'][label]['precision']:.4f}",
            f"{selected['classification_report'][label]['recall']:.4f}",
            f"{selected['classification_report'][label]['f1-score']:.4f}",
            str(int(selected["classification_report"][label]["support"])),
        ]
        for label in metrics["class_order"]
    )
    story.append(_table(class_rows, [1.45 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch, 1.0 * inch], styles))
    story.append(Spacer(1, 7))
    callout = Table(
        [[_p(
            f"Interpretation: the selected classifier correctly labeled {selected['accuracy']:.2%} of this synthetic test set; macro F1 summarizes balanced class performance and is {selected['macro_f1']:.4f}. These are prototype results, not validated production estimates.",
            styles["callout"],
        )]],
        colWidths=[usable_width],
    )
    callout.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE),
                ("BOX", (0, 0), (-1, -1), 0.8, TEAL),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(callout)

    # Page 3: Round 3 insights, prototype, and recommendations.
    story.append(PageBreak())
    heading("h1", "Round 3 - Intelligence Layer, Visualization & Prototype")
    heading("h2", "1. Visualization")
    chart_rows = [["Generated chart", "Saved file"]]
    chart_rows.extend(
        [CHART_TITLES.get(name, name), name]
        for name in evidence["chart_files"]
    )
    story.append(
        _table(
            chart_rows,
            [3.55 * inch, usable_width - 3.55 * inch],
            styles,
        )
    )
    body(
        f"The Excel workbook contains a Summary Dashboard, {len(evidence['workbook_sheets']) - 1} analysis/data sheets, {evidence['workbook_tables']} formatted tables, and {evidence['workbook_charts']} embedded charts. The implemented Streamlit app provides age-band and region filters and shows engagement, channels, account activity, complaints, and service trends."
    )
    # Use two verified chart files as small visual examples.
    images = []
    for filename in ("channel_usage.png", "regional_complaints.png"):
        image_path = ROOT / "outputs" / "charts" / filename
        if image_path.exists():
            image = Image(str(image_path), width=3.15 * inch, height=1.72 * inch)
            images.append(image)
    if len(images) == 2:
        image_table = Table([[images[0], images[1]]], colWidths=[3.48 * inch, 3.48 * inch])
        image_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 2),
                ]
            )
        )
        story.append(image_table)
        story.append(Spacer(1, 3))
    heading("h2", "2. Key Insights")
    term = findings["account_activity_q1_q4"]["Term Deposit"]
    repeats = findings["repeated_requests_by_service_type"]
    body(
        f"Mobile led channel usage ({findings['top_channel_uses']:,} uses), and North had the most complaints ({findings['regional_complaint_counts']['North']:,})."
    )
    body(
        f"Term Deposit activity declined {abs(term['change_percent']):.2f}% from Q1 ({term['q1_transactions']:,}) to Q4 ({term['q4_transactions']:,}) 2025. Card support and Payments tied for most repeated request types at {repeats['Card support']} customer/service-type pairs each."
    )
    heading("h2", "3. Prototype / Intelligence Layer")
    body(
        "Python data preparation and Scikit-learn pipelines create and evaluate customer engagement classifications. The working Streamlit application presents filtered customer, channel, account, complaint, and service insights. The Excel workbook is a separate analysis deliverable; the PowerPoint pitch is in the presentation folder."
    )
    heading("h2", "4. Recommendations")
    body(
        f"1. Pilot assisted digital onboarding for customers aged 60+ (average mobile {findings['channel_average_uses_by_age_band']['60+']['mobile_uses']:.2f}, internet {findings['channel_average_uses_by_age_band']['60+']['internet_uses']:.2f} uses, the lowest combined age-band average); measure adoption before scaling."
    )
    body(
        f"2. Review complaint drivers and resolution in North ({findings['regional_complaint_counts']['North']:,} complaints), investigate Card support and Payments repeat contacts ({repeats['Card support']} each), and test targeted Term Deposit follow-up after the measured activity decline."
    )

    # Page 4: Excel evidence, business pitch, and final submission.
    story.append(PageBreak())
    heading("h1", "Excel Analysis, Business Pitch & Final Submission")
    heading("h2", "Round 3 - Data Analysis Using Excel: Evidence Checklist")
    body(
        "Workbook inspected: `excel/bank_customer_insight_dashboard.xlsx`. Status describes the workbook itself, not Python charts or reports.",
        "small",
    )
    checklist_rows = [["Assessment item", "Status", "Verified workbook evidence"]]
    checklist_rows.extend([name, status, detail] for name, status, detail in _status(evidence))
    checklist_table = _table(
        checklist_rows,
        [1.82 * inch, 0.78 * inch, usable_width - 2.6 * inch],
        styles,
    )
    checklist_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (1, 1), (1, -1), colors.white),
                ("TEXTCOLOR", (1, 1), (1, -1), NAVY),
            ]
        )
    )
    story.append(checklist_table)
    body(
        "PivotTables/PivotCharts were not generated natively. The workbook includes desktop Excel PivotTable instructions; create these manually if the assessment requires native pivots.",
        "small",
    )
    heading("h2", "Round 4 - Pitching Business Insights")
    pitch_text = _pitch(findings)
    pitch_box = Table(
        [[_p(pitch_text, styles["body"])]],
        colWidths=[usable_width],
    )
    pitch_box.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), PALE),
                ("BOX", (0, 0), (-1, -1), 0.8, NAVY),
                ("LEFTPADDING", (0, 0), (-1, -1), 9),
                ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(pitch_box)
    words = len(pitch_text.split())
    body(f"Pitch length: {words} words.", "small")
    heading("h2", "Final Submission")
    repository_line = evidence["repository_url"] or "____________________________________________"
    body(f"GitHub Repository Link: {repository_line}")
    body(
        f"Overall Result / Conclusion: The synthetic-data workflow links customer, channel, account, transaction, and service patterns and compares three engagement classifiers. Mobile led channel use, North had the highest complaint count, and Logistic Regression achieved {selected['accuracy']:.2%} test accuracy with macro F1 {selected['macro_f1']:.4f}. Validate insights against governed real data before operational use."
    )

    document = SimpleDocTemplate(
        str(PDF_PATH),
        pagesize=letter,
        rightMargin=0.5 * inch,
        leftMargin=0.5 * inch,
        topMargin=0.43 * inch,
        bottomMargin=0.62 * inch,
        title="CH08 Bank Customer Insight Assessment Answers",
        author="Vissvapriya P",
        subject="DS Sprint Skill-Based Assessment 2 - Use Case 08",
    )
    document.build(story, onFirstPage=_footer, onLaterPages=_footer)


def main() -> None:
    REPORTS.mkdir(parents=True, exist_ok=True)
    evidence = load_evidence()
    markdown = build_markdown(evidence)
    MARKDOWN_PATH.write_text(markdown, encoding="utf-8")
    build_pdf(evidence)
    print(f"Markdown: {MARKDOWN_PATH}")
    print(f"PDF: {PDF_PATH}")
    print(f"Verified source shapes: {json.dumps(evidence['datasets'], sort_keys=True)}")
    print(
        "Excel checks: "
        f"{evidence['workbook_tables']} tables, {evidence['workbook_charts']} charts, "
        f"{evidence['workbook_formulas']} formulas, {evidence['workbook_pivots']} PivotTables."
    )


if __name__ == "__main__":
    main()
