"""Build the Excel dashboard, handwritable assessment answers, and pitch deck."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]


def _read_json(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


def _recommended_actions(findings: dict[str, object]) -> list[str]:
    age_usage = findings["channel_average_uses_by_age_band"]
    digital_by_age = {
        age: values["mobile_uses"] + values["internet_uses"]
        for age, values in age_usage.items()
    }
    least_digital_age = min(digital_by_age, key=digital_by_age.get)
    region = findings["highest_complaint_region"]
    repeat_type = findings["most_repeated_service_type"]
    declining = findings["declining_account_types"]
    decline_note = (
        f"Review retention journeys for {', '.join(declining)} accounts, which recorded lower Q4 than Q1 activity."
        if declining
        else "Monitor Q1-to-Q4 account activity and test retention prompts for account types as activity changes."
    )
    return [
        (
            f"Pilot assisted digital onboarding for customers aged {least_digital_age}, "
            f"the age band with the lowest combined average mobile and internet use; "
            "measure adoption and repeat use before scaling."
        ),
        (
            f"Investigate complaint drivers in {region} and redesign follow-up for "
            f"{repeat_type}, the service type with the most customers making repeat requests. "
            f"{decline_note}"
        ),
    ]


def _format_workbook(path: Path, findings: dict[str, object], ml: dict[str, object]) -> None:
    processed = ROOT / "data" / "processed"
    feature_frame = pd.read_csv(processed / "customer_insight_features.csv")
    target = pd.read_csv(processed / "customer_engagement_target.csv")
    sheets: dict[str, pd.DataFrame] = {
        "Customer Data": pd.read_csv(processed / "customers_clean.csv").merge(
            target, on="customer_id", validate="one_to_one"
        ),
        "Account Data": pd.read_csv(processed / "accounts_clean.csv"),
        "Transactions": pd.read_csv(processed / "transactions_clean.csv"),
        "Channel Analysis": pd.read_csv(processed / "channel_analysis.csv"),
        "Customer Segments": feature_frame.merge(
            target, on="customer_id", validate="one_to_one"
        ),
        "Account Activity": pd.read_csv(processed / "account_activity.csv"),
        "Complaint Analysis": pd.read_csv(processed / "complaint_analysis.csv"),
        "Balance Analysis": pd.read_csv(processed / "balance_analysis.csv"),
        "Service Trends": pd.read_csv(processed / "service_trends.csv"),
    }
    dashboard = pd.DataFrame(
        [
            ["TrustWave Bank | Customer Insight", "Synthetic-data analysis; results are illustrative only."],
            ["Customers", "=COUNTA('Customer Data'!A2:A1201)"],
            ["Transactions after deduplication", "=COUNTA(Transactions!A2:A1048576)"],
            ["Most-used channel", findings["top_channel"].title()],
            ["Top channel uses", findings["top_channel_uses"]],
            ["Highest-complaint region", findings["highest_complaint_region"]],
            ["Most repeated service type", findings["most_repeated_service_type"]],
            ["Selected model", ml["selected_model"]],
            [
                "Selected model test accuracy",
                ml["models"][ml["selected_model"]]["accuracy"],
            ],
            [
                "Selected model test macro F1",
                ml["models"][ml["selected_model"]]["macro_f1"],
            ],
            ["How to create native PivotTables", "Select a formatted data table, then use Insert > PivotTable in desktop Excel."],
            ["Pivot suggestion 1", "Transactions: Rows = account_type; Columns = transaction_date (group by quarters); Values = Count of transaction_id."],
            ["Pivot suggestion 2", "Service: Rows = region; Columns = is_complaint; Values = Count of service_id."],
        ],
        columns=["Metric", "Value / instruction"],
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        dashboard.to_excel(writer, sheet_name="Summary Dashboard", index=False)
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)

    workbook = load_workbook(path)
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        sheet.sheet_view.showGridLines = False
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="163A5F")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        sheet.row_dimensions[1].height = 28
        for column_cells in sheet.columns:
            column_letter = column_cells[0].column_letter
            sample_values = [str(cell.value or "") for cell in column_cells[:200]]
            width = min(48, max(12, max(map(len, sample_values), default=12) + 2))
            sheet.column_dimensions[column_letter].width = width
        if sheet.max_row > 1 and sheet.title != "Summary Dashboard":
            safe_name = "".join(ch for ch in sheet.title if ch.isalnum())[:20]
            table = Table(displayName=f"tbl{safe_name}", ref=sheet.dimensions)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            sheet.add_table(table)

    dashboard_sheet = workbook["Summary Dashboard"]
    dashboard_sheet.freeze_panes = "A2"
    for row in range(2, dashboard_sheet.max_row + 1):
        dashboard_sheet.cell(row=row, column=1).font = Font(bold=True, color="163A5F")
        dashboard_sheet.cell(row=row, column=2).alignment = Alignment(wrap_text=True)
    dashboard_sheet.column_dimensions["A"].width = 38
    dashboard_sheet.column_dimensions["B"].width = 100

    channel_sheet = workbook["Channel Analysis"]
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Total Usage by Channel"
    chart.y_axis.title = "Channel"
    chart.x_axis.title = "Uses"
    chart.add_data(Reference(channel_sheet, min_col=2, min_row=1, max_row=channel_sheet.max_row), titles_from_data=True)
    chart.set_categories(Reference(channel_sheet, min_col=1, min_row=2, max_row=channel_sheet.max_row))
    chart.height, chart.width = 7, 12
    dashboard_sheet.add_chart(chart, "D2")

    activity_sheet = workbook["Account Activity"]
    chart = BarChart()
    chart.type = "col"
    chart.style = 11
    chart.title = "Account Activity: Q1 vs Q4"
    chart.y_axis.title = "Transactions"
    chart.add_data(Reference(activity_sheet, min_col=2, max_col=3, min_row=1, max_row=activity_sheet.max_row), titles_from_data=True)
    chart.set_categories(Reference(activity_sheet, min_col=1, min_row=2, max_row=activity_sheet.max_row))
    chart.height, chart.width = 7, 12
    dashboard_sheet.add_chart(chart, "D17")
    workbook.save(path)


def _add_slide_title(slide, title: str, subtitle: str = "Synthetic-data findings | TrustWave Bank") -> None:
    heading = slide.shapes.title
    heading.text = title
    heading.text_frame.paragraphs[0].font.size = Pt(28)
    heading.text_frame.paragraphs[0].font.bold = True
    heading.text_frame.paragraphs[0].font.color.rgb = RGBColor(22, 58, 95)
    textbox = slide.shapes.add_textbox(Inches(0.55), Inches(7.08), Inches(12.1), Inches(0.25))
    paragraph = textbox.text_frame.paragraphs[0]
    paragraph.text = subtitle
    paragraph.font.size = Pt(9)
    paragraph.font.color.rgb = RGBColor(90, 100, 110)


def _add_bullets(slide, bullets: list[str], x=0.75, y=1.5, width=11.7, height=4.9) -> None:
    textbox = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = textbox.text_frame
    frame.word_wrap = True
    for index, item in enumerate(bullets):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.text = item
        paragraph.level = 0
        paragraph.font.size = Pt(21)
        paragraph.space_after = Pt(14)


def _make_pitch(path: Path, findings: dict[str, object], ml: dict[str, object]) -> None:
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    layout = presentation.slide_layouts[5]

    slide = presentation.slides.add_slide(presentation.slide_layouts[0])
    slide.shapes.title.text = "Bank Customer Insight"
    slide.placeholders[1].text = (
        "DS Sprint Skill-Based Assessment 2 | Use Case 08\n"
        "TrustWave Bank | Synthetic-data study"
    )
    slide.shapes.title.text_frame.paragraphs[0].font.color.rgb = RGBColor(22, 58, 95)

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Business problem and objectives")
    _add_bullets(
        slide,
        [
            "Business question: How do customers use banking services, and which segments show changing engagement?",
            "Understand channel preferences, account activity, balance differences, complaints, and repeat service needs.",
            "Build an interpretable customer-level engagement prototype to inform targeted service decisions.",
        ],
    )

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Synthetic dataset overview")
    _add_bullets(
        slide,
        [
            f"{findings['source_record_counts']['customers']:,} customers across five regions and four age bands.",
            f"{findings['source_record_counts']['accounts']:,} accounts; {findings['source_record_counts']['transactions']:,} raw transaction rows including injected duplicates.",
            f"{findings['source_record_counts']['channels']:,} channel profiles and {findings['source_record_counts']['service']:,} service records.",
            "Fixed seed 20261013. All records and results are synthetic; they do not describe real bank customers.",
        ],
    )

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Data preparation")
    _add_bullets(
        slide,
        [
            f"Removed {findings['duplicate_transaction_rows_removed']:,} repeated transaction IDs; unique transaction IDs preserve similar legitimate payments.",
            f"Standardized inconsistent account labels and filled {findings['missing_channel_values_filled']:,} missing channel cells using age-band medians.",
            f"IQR bounds: {findings['transaction_amount_iqr_bounds']['lower']:.2f} to {findings['transaction_amount_iqr_bounds']['upper']:.2f}; flagged {findings['transaction_amount_outliers_flagged_and_capped']:,} and capped extreme amounts.",
            "Filled incomplete service categories/statuses with explicit labels and imputed missing resolution time; saved before/after quality counts.",
        ],
    )

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "EDA: channels and customer balances")
    chart = slide.shapes.add_picture(
        str(ROOT / "outputs" / "charts" / "channel_usage.png"),
        Inches(0.6), Inches(1.45), width=Inches(6.0),
    )
    _add_bullets(
        slide,
        [
            f"{findings['top_channel'].title()} is the most-used channel ({findings['top_channel_uses']:,} reported uses).",
            f"Highest complaint count: {findings['highest_complaint_region']} region.",
            "Mean total balance by segment: "
            + "; ".join(
                f"{segment} ${value:,.0f}"
                for segment, value in findings["average_balance_by_engagement_segment"].items()
            )
            + ".",
        ],
        x=6.8, y=1.55, width=5.7, height=4.8,
    )

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Age, account activity, and service")
    slide.shapes.add_picture(
        str(ROOT / "outputs" / "charts" / "age_group_comparison.png"),
        Inches(0.55), Inches(1.45), width=Inches(5.9),
    )
    slide.shapes.add_picture(
        str(ROOT / "outputs" / "charts" / "account_activity.png"),
        Inches(6.8), Inches(1.45), width=Inches(5.9),
    )
    _add_bullets(
        slide,
        [
            f"Account types with lower Q4 than Q1 transaction activity: {', '.join(findings['declining_account_types']) or 'none in this synthetic sample'}.",
            f"Most repeated service type: {findings['most_repeated_service_type']}.",
        ],
        x=0.8, y=5.75, width=11.8, height=1.05,
    )

    selected = ml["models"][ml["selected_model"]]
    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Machine-learning results")
    _add_bullets(
        slide,
        [
            f"Compared majority-class DummyClassifier, Logistic Regression, and Random Forest on a stratified 75/25 split.",
            f"Selected {ml['selected_model']} by test macro F1 (accuracy breaks ties): accuracy {selected['accuracy']:.3f}, macro F1 {selected['macro_f1']:.3f}.",
            f"Test set: {ml['test_records']:,} customers. Full class report and confusion matrix are saved under outputs/metrics/.",
            "Target is simulated from a latent engagement propensity; the latent value is not a predictor feature.",
        ],
    )

    slide = presentation.slides.add_slide(layout)
    _add_slide_title(slide, "Evidence-based recommendations and conclusion")
    _add_bullets(slide, _recommended_actions(findings), height=5.5)
    textbox = slide.shapes.add_textbox(Inches(0.75), Inches(6.1), Inches(11.7), Inches(0.55))
    p = textbox.text_frame.paragraphs[0]
    p.text = "Treat results as a synthetic prototype; validate any intervention using real, governed bank data."
    p.font.size = Pt(15)
    p.font.italic = True
    p.alignment = PP_ALIGN.CENTER

    path.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(path)


def _write_assessment_answers(
    findings: dict[str, object], ml: dict[str, object]
) -> None:
    selected = ml["models"][ml["selected_model"]]
    quality = findings["data_quality"]["before"]
    source_lines = "\n".join(
        f"   - `{name}.csv`: {details['records']:,} records; {details['attribute_count']} attributes."
        for name, details in quality.items()
    )
    answers = f"""# DS Sprint Skill-Based Assessment 2 — Use Case 08

- **Name:** Vissvapriya P
- **Register Number:**
- **Class:** III CSE D
- **Challenge ID:** CH08
- **Title:** Bank Customer Insight
- **Data note:** All records, findings, and model results below are synthetic (seed {findings['seed']}).

## Round 1

### Problem Statement (2–3 lines)
Understand how TrustWave Bank customers use channels and accounts, and identify segments whose engagement is changing.
Use customer, transaction, and service patterns to suggest practical improvements to digital access and customer support.

### Dataset Details (names, source, actual records, attributes)
Source: generated synthetic data; the five CSV files are in `data/raw/`.
{source_lines}

### Data Preprocessing
Filled {findings['missing_channel_values_filled']:,} missing channel values with age-band medians and completed missing service labels/times.
Standardized account names, removed {findings['duplicate_transaction_rows_removed']:,} duplicate transaction IDs, and flagged/capped {findings['transaction_amount_outliers_flagged_and_capped']:,} IQR outliers. Saved cleaned tables in `data/processed/`.

### Two EDA Findings
1. **Channel use:** {findings['top_channel'].title()} was most used, with {findings['top_channel_uses']:,} reported uses.
2. **Regional complaints:** {findings['highest_complaint_region']} had the highest complaint count ({findings['regional_complaint_counts'][findings['highest_complaint_region']]:,}).

Six business-question results: age-band channel-use averages, Q1-to-Q4 account activity, regional complaint counts, segment balance averages, and repeated requests by service type are in `outputs/metrics/round1_findings.json` and the project report.

## Round 2

### Selected Features
Age band, region, relationship months, total account balance, account count, transaction count, transaction frequency per month, branch/ATM/mobile/internet uses, complaint count, and service request count.

### ML Algorithm Used
DummyClassifier baseline, Logistic Regression, and Random Forest Classifier.

### Justification
Compared all three with the same stratified 75/25 split and preprocessing pipeline. Selected **{ml['selected_model']}** for the highest test macro F1 (accuracy breaks ties), to account for performance across all three classes. The synthetic target is based on a simulated latent propensity not included among the predictors.

### Model Performance with actual accuracy and F1-score
Selected model test accuracy: **{selected['accuracy']:.4f}**. Test macro F1-score: **{selected['macro_f1']:.4f}** ({ml['test_records']:,} test customers). Full report and confusion matrix are in `outputs/metrics/`.

## Round 3

### Visualization
Created channel, age-band, account activity, regional complaint, balance, monthly transaction, and service-request charts. Excel workbook includes formatted tables, formulas, and charts. Use Excel's **Insert > PivotTable** on a data table to create native PivotTables.

### Two Key Insights
1. {findings['top_channel'].title()} leads channel use ({findings['top_channel_uses']:,} uses); channel patterns differ by age band.
2. {findings['highest_complaint_region']} has the largest complaint count; {findings['most_repeated_service_type']} is the most repeated request type.

### Prototype / Intelligence Layer
The Streamlit app displays customer engagement segments, channel use, account activity, complaints, and service trends. A Scikit-learn pipeline compares a baseline and two classifiers.

### Two Recommendations
{chr(10).join(f'{index}. {action}' for index, action in enumerate(_recommended_actions(findings), start=1))}

## Final Submission

- **GitHub Repository Link:**
- **Overall Result / Conclusion:** This synthetic analysis demonstrates a repeatable way to connect banking activity, channels, and service needs. Validate the patterns with governed real data before making operational decisions.
"""
    (ROOT / "reports" / "assessment_sheet_answers.md").write_text(answers, encoding="utf-8")


def _write_project_report(findings: dict[str, object], ml: dict[str, object]) -> None:
    selected = ml["models"][ml["selected_model"]]
    before = findings["data_quality"]["before"]
    after = findings["data_quality"]["after"]
    quality_rows = "\n".join(
        f"| {name} | {before[name]['records']:,} | {after[name]['records']:,} | "
        f"{before[name]['missing_cells']:,} | {after[name]['missing_cells']:,} | "
        f"{before[name]['duplicate_rows']:,} | {after[name]['duplicate_rows']:,} |"
        for name in before
    )
    age_rows = "\n".join(
        f"| {age} | " + " | ".join(f"{values[channel + '_uses']:.2f}" for channel in ["branch", "atm", "mobile", "internet"]) + " |"
        for age, values in findings["channel_average_uses_by_age_band"].items()
    )
    acct_rows = "\n".join(
        f"| {name} | {values['q1_transactions']:,} | {values['q4_transactions']:,} | {values['change_percent']:.2f}% |"
        for name, values in findings["account_activity_q1_q4"].items()
    )
    region_rows = "\n".join(
        f"| {region} | {count:,} |" for region, count in findings["regional_complaint_counts"].items()
    )
    balance_rows = "\n".join(
        f"| {segment} | ${value:,.2f} |"
        for segment, value in findings["average_balance_by_engagement_segment"].items()
    )
    service_rows = "\n".join(
        f"| {name} | {count:,} |"
        for name, count in findings["repeated_requests_by_service_type"].items()
    ) or "| No repeated categories | 0 |"
    model_rows = "\n".join(
        f"| {name} | {values['accuracy']:.4f} | {values['macro_f1']:.4f} |"
        for name, values in ml["models"].items()
    )
    report = f"""# Bank Customer Insight — Project Report

**Organization:** TrustWave Bank  
**Sector:** Banking  
**Business question:** How do customers use banking services, and which segments show changing engagement?  
**Data disclaimer:** Every record and result in this project is generated synthetic data (seed {findings['seed']}), not real customer information.

## Objective and scope

Explore five linked customer/account/transaction/channel/service tables; answer all six assessment questions; create a leakage-aware engagement classification prototype; deliver a workbook, interactive dashboard, and pitch deck.

## Dataset and data-quality review

Raw source record counts: {json.dumps(findings['source_record_counts'], sort_keys=True)}.

| Source table | Before rows | After rows | Before missing cells | After missing cells | Before exact duplicates | After exact duplicates |
|---|---:|---:|---:|---:|---:|---:|
{quality_rows}

Columns, source attributes, actual dtypes, per-column missing-value counts, and duplicate counts are in `outputs/metrics/data_quality_report.json`. Transactions are de-duplicated by unique `transaction_id`, so similar but separately identified transactions remain. Account labels are normalized into four canonical values. Channel gaps use age-band medians. Incomplete service labels receive explicit `Unspecified`/`Unknown` values; resolution days use the observed median. Transaction amounts outside the IQR fences ({findings['transaction_amount_iqr_bounds']['lower']:.2f}, {findings['transaction_amount_iqr_bounds']['upper']:.2f}) are flagged and capped while the original value is retained in `amount_original`.

## Round 1 — Answers to the six business questions

### 1. Which channels are used most?

{findings['top_channel'].title()} leads with {findings['top_channel_uses']:,} recorded channel uses. Totals: {json.dumps(findings['channel_totals'], sort_keys=True)}.

### 2. How does channel use differ by age band?

Average recorded uses per customer:

| Age band | Branch | ATM | Mobile | Internet |
|---|---:|---:|---:|---:|
{age_rows}

### 3. Which account types show declining activity?

Q1 (January–March) and Q4 (October–December) 2025 counts:

| Account type | Q1 transactions | Q4 transactions | Q1-to-Q4 change |
|---|---:|---:|---:|
{acct_rows}

Account types with lower Q4 than Q1: {', '.join(findings['declining_account_types']) or 'none in this synthetic sample'}.

### 4. Which regions generate more complaints?

| Region | Complaint records |
|---|---:|
{region_rows}

Highest count: {findings['highest_complaint_region']}.

### 5. How do balances differ by segment?

Mean total customer account balance by simulated engagement segment:

| Segment | Mean balance |
|---|---:|
{balance_rows}

### 6. Which services drive repeated requests?

Counts of customer/service-type pairs with more than one request:

| Service type | Customers with repeat requests |
|---|---:|
{service_rows}

Most repeated category: {findings['most_repeated_service_type']}.

## Round 2 — Features and model evaluation

The target classes are tertiles of a simulated latent engagement propensity. The latent value itself is not a feature; it influences generated behavioral signals. Predictors include age band, region, relationship length, total balance, account count, transaction count and monthly frequency, channel counts, complaints, and total service requests. Numeric features are standardized, categorical features one-hot encoded inside a Scikit-learn pipeline. A stratified 75/25 holdout uses random state 42.

| Model | Test accuracy | Test macro F1 |
|---|---:|---:|
{model_rows}

Selected model: **{ml['selected_model']}** (best macro F1; accuracy breaks ties). Test accuracy {selected['accuracy']:.4f}; macro F1 {selected['macro_f1']:.4f}. The model report, test predictions, and confusion matrix are saved under `outputs/metrics/`. These results measure prediction of a synthetic target and should not be interpreted as real-world bank-model performance.

## Round 3 — Deliverables

- Seven EDA charts in `outputs/charts/`.
- `excel/bank_customer_insight_dashboard.xlsx`: ten analysis/dashboard sheets, formatted tables, formulas, and embedded charts.
- `app.py`: interactive Streamlit prototype. OpenPyXL cannot author native Excel PivotTables; steps and field suggestions are included on the workbook dashboard.
- `presentation/bank_customer_insight_pitch.pptx`: concise evidence-based presentation.

## Evidence-based recommendations

{chr(10).join(f'{i}. {action}' for i, action in enumerate(_recommended_actions(findings), start=1))}

## Limitations and conclusion

Synthetic distributions are intentionally designed for demonstration. The target represents simulated latent propensity, not observed customer outcomes. Before operational use, validate with governed real data, suitable privacy and fairness checks, temporal validation, and business-owner review.
"""
    (ROOT / "reports" / "project_report.md").write_text(report, encoding="utf-8")


def run_round3() -> dict[str, str]:
    """Write analysis workbook, assessment answers, report, and PowerPoint."""
    findings = _read_json(ROOT / "outputs" / "metrics" / "round1_findings.json")
    ml = _read_json(ROOT / "outputs" / "metrics" / "model_metrics.json")
    excel_path = ROOT / "excel" / "bank_customer_insight_dashboard.xlsx"
    pptx_path = ROOT / "presentation" / "bank_customer_insight_pitch.pptx"
    _format_workbook(excel_path, findings, ml)
    _write_assessment_answers(findings, ml)
    _write_project_report(findings, ml)
    _make_pitch(pptx_path, findings, ml)
    return {
        "excel": str(excel_path),
        "assessment_answers": str(ROOT / "reports" / "assessment_sheet_answers.md"),
        "project_report": str(ROOT / "reports" / "project_report.md"),
        "presentation": str(pptx_path),
    }


if __name__ == "__main__":
    print(json.dumps(run_round3(), indent=2))
