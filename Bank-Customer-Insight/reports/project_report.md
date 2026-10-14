# Bank Customer Insight — Project Report

**Organization:** TrustWave Bank  
**Sector:** Banking  
**Business question:** How do customers use banking services, and which segments show changing engagement?  
**Data disclaimer:** Every record and result in this project is generated synthetic data (seed 20261013), not real customer information.

## Objective and scope

Explore five linked customer/account/transaction/channel/service tables; answer all six assessment questions; create a leakage-aware engagement classification prototype; deliver a workbook, interactive dashboard, and pitch deck.

## Dataset and data-quality review

Raw source record counts: {"accounts": 1907, "channels": 1200, "customers": 1200, "service": 1971, "transactions": 48169}.

| Source table | Before rows | After rows | Before missing cells | After missing cells | Before exact duplicates | After exact duplicates |
|---|---:|---:|---:|---:|---:|---:|
| customers | 1,200 | 1,200 | 0 | 0 | 0 | 0 |
| accounts | 1,907 | 1,907 | 0 | 0 | 0 | 0 |
| transactions | 48,169 | 47,693 | 0 | 0 | 476 | 0 |
| channels | 1,200 | 1,200 | 277 | 0 | 0 | 0 |
| service | 1,971 | 1,971 | 370 | 0 | 0 | 0 |

Columns, source attributes, actual dtypes, per-column missing-value counts, and duplicate counts are in `outputs/metrics/data_quality_report.json`. Transactions are de-duplicated by unique `transaction_id`, so similar but separately identified transactions remain. Account labels are normalized into four canonical values. Channel gaps use age-band medians. Incomplete service labels receive explicit `Unspecified`/`Unknown` values; resolution days use the observed median. Transaction amounts outside the IQR fences (0.00, 279.81) are flagged and capped while the original value is retained in `amount_original`.

## Round 1 — Answers to the six business questions

### 1. Which channels are used most?

Mobile leads with 5,286 recorded channel uses. Totals: {"atm": 4312, "branch": 2963, "internet": 3255, "mobile": 5286}.

### 2. How does channel use differ by age band?

Average recorded uses per customer:

| Age band | Branch | ATM | Mobile | Internet |
|---|---:|---:|---:|---:|
| 18-29 | 2.29 | 3.46 | 6.31 | 3.99 |
| 30-44 | 2.54 | 3.39 | 5.23 | 3.13 |
| 45-59 | 2.36 | 3.77 | 3.52 | 2.21 |
| 60+ | 2.68 | 3.84 | 2.16 | 1.31 |

### 3. Which account types show declining activity?

Q1 (January–March) and Q4 (October–December) 2025 counts:

| Account type | Q1 transactions | Q4 transactions | Q1-to-Q4 change |
|---|---:|---:|---:|
| Checking | 2,957 | 3,013 | 1.89% |
| Credit Card | 2,658 | 3,446 | 29.65% |
| Savings | 2,578 | 2,613 | 1.36% |
| Term Deposit | 3,620 | 2,860 | -20.99% |

Account types with lower Q4 than Q1: Term Deposit.

### 4. Which regions generate more complaints?

| Region | Complaint records |
|---|---:|
| North | 169 |
| West | 158 |
| South | 126 |
| East | 100 |
| Central | 91 |

Highest count: North.

### 5. How do balances differ by segment?

Mean total customer account balance by simulated engagement segment:

| Segment | Mean balance |
|---|---:|
| Low | $14,920.88 |
| Medium | $16,442.77 |
| High | $18,415.64 |

### 6. Which services drive repeated requests?

Counts of customer/service-type pairs with more than one request:

| Service type | Customers with repeat requests |
|---|---:|
| Card support | 60 |
| Payments | 60 |
| Digital access | 54 |
| Account inquiry | 47 |
| Loan service | 32 |
| Unspecified | 1 |

Most repeated category: Card support.

## Round 2 — Features and model evaluation

The target classes are tertiles of a simulated latent engagement propensity. The latent value itself is not a feature; it influences generated behavioral signals. Predictors include age band, region, relationship length, total balance, account count, transaction count and monthly frequency, channel counts, complaints, and total service requests. Numeric features are standardized, categorical features one-hot encoded inside a Scikit-learn pipeline. A stratified 75/25 holdout uses random state 42.

| Model | Test accuracy | Test macro F1 |
|---|---:|---:|
| DummyClassifier | 0.3333 | 0.1667 |
| Logistic Regression | 0.7033 | 0.6991 |
| Random Forest | 0.6967 | 0.6945 |

Selected model: **Logistic Regression** (best macro F1; accuracy breaks ties). Test accuracy 0.7033; macro F1 0.6991. The model report, test predictions, and confusion matrix are saved under `outputs/metrics/`. These results measure prediction of a synthetic target and should not be interpreted as real-world bank-model performance.

## Round 3 — Deliverables

- Seven EDA charts in `outputs/charts/`.
- `excel/bank_customer_insight_dashboard.xlsx`: ten analysis/dashboard sheets, formatted tables, formulas, and embedded charts.
- `app.py`: interactive Streamlit prototype. OpenPyXL cannot author native Excel PivotTables; steps and field suggestions are included on the workbook dashboard.
- `presentation/bank_customer_insight_pitch.pptx`: concise evidence-based presentation.

## Evidence-based recommendations

1. Pilot assisted digital onboarding for customers aged 60+, the age band with the lowest combined average mobile and internet use; measure adoption and repeat use before scaling.
2. Investigate complaint drivers in North and redesign follow-up for Card support, the service type with the most customers making repeat requests. Review retention journeys for Term Deposit accounts, which recorded lower Q4 than Q1 activity.

## Limitations and conclusion

Synthetic distributions are intentionally designed for demonstration. The target represents simulated latent propensity, not observed customer outcomes. Before operational use, validate with governed real data, suitable privacy and fairness checks, temporal validation, and business-owner review.
