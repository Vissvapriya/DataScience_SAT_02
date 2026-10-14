"""Leakage-safe customer feature engineering and model comparison."""

from __future__ import annotations

import json
from pathlib import Path

import joblib
import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.dummy import DummyClassifier
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import (
    accuracy_score,
    classification_report,
    confusion_matrix,
    f1_score,
)
from sklearn.model_selection import train_test_split
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

ROOT = Path(__file__).resolve().parents[1]
CLASS_ORDER = ["Low", "Medium", "High"]


def build_customer_features() -> pd.DataFrame:
    """Aggregate the cleaned source tables to one row per customer."""
    processed = ROOT / "data" / "processed"
    customers = pd.read_csv(processed / "customers_clean.csv")
    accounts = pd.read_csv(processed / "accounts_clean.csv")
    transactions = pd.read_csv(processed / "transactions_clean.csv")
    channels = pd.read_csv(processed / "channels_clean.csv")
    service = pd.read_csv(processed / "service_clean.csv")

    balances = accounts.groupby("customer_id")["balance"].agg(
        account_balance="sum", account_count="count"
    )
    transaction_summary = transactions.groupby("account_id").size().rename("account_transaction_count")
    account_customer = accounts[["account_id", "customer_id"]].merge(
        transaction_summary, on="account_id", how="left", validate="one_to_one"
    )
    transaction_counts = (
        account_customer.groupby("customer_id")["account_transaction_count"]
        .sum()
        .rename("transaction_count")
    )
    complaint_counts = (
        service.loc[service["is_complaint"].astype(bool)]
        .groupby("customer_id")
        .size()
        .rename("complaint_count")
    )
    service_counts = service.groupby("customer_id").size().rename("service_request_count")
    features = customers.set_index("customer_id").join(
        [balances, transaction_counts, complaint_counts, service_counts]
    )
    features = features.join(channels.set_index("customer_id"), validate="one_to_one")
    features = features.drop(columns=["customer_id"], errors="ignore")
    features["account_balance"] = features["account_balance"].fillna(0)
    features["account_count"] = features["account_count"].fillna(0).astype(int)
    features["transaction_count"] = features["transaction_count"].fillna(0).astype(int)
    features["complaint_count"] = features["complaint_count"].fillna(0).astype(int)
    features["service_request_count"] = features["service_request_count"].fillna(0).astype(int)
    features["transaction_frequency_per_month"] = features["transaction_count"] / 12
    features.index.name = "customer_id"
    result = features.reset_index()
    result.to_csv(processed / "customer_insight_features.csv", index=False)
    return result


def run_round2() -> dict[str, object]:
    """Train a baseline and two classifiers, then save actual test results."""
    features = build_customer_features()
    target = pd.read_csv(ROOT / "data" / "processed" / "customer_engagement_target.csv")
    dataset = features.merge(target, on="customer_id", validate="one_to_one")
    y = dataset.pop("engagement_segment")
    identifiers = dataset.pop("customer_id")
    x = dataset

    categorical = ["age_band", "region"]
    numeric = [name for name in x.columns if name not in categorical]
    preprocessor = ColumnTransformer(
        transformers=[
            ("numeric", StandardScaler(), numeric),
            ("categorical", OneHotEncoder(handle_unknown="ignore"), categorical),
        ]
    )
    x_train, x_test, y_train, y_test, id_train, id_test = train_test_split(
        x,
        y,
        identifiers,
        test_size=0.25,
        random_state=42,
        stratify=y,
    )
    candidates = {
        "DummyClassifier": DummyClassifier(strategy="most_frequent"),
        "Logistic Regression": LogisticRegression(max_iter=1500, class_weight="balanced"),
        "Random Forest": RandomForestClassifier(
            n_estimators=250,
            min_samples_leaf=3,
            class_weight="balanced",
            random_state=42,
            n_jobs=-1,
        ),
    }
    fitted: dict[str, Pipeline] = {}
    metrics: dict[str, dict[str, object]] = {}
    prediction_rows = []
    for name, classifier in candidates.items():
        model = Pipeline(
            steps=[("preprocessing", preprocessor), ("classifier", classifier)]
        )
        model.fit(x_train, y_train)
        predictions = model.predict(x_test)
        fitted[name] = model
        metrics[name] = {
            "accuracy": float(accuracy_score(y_test, predictions)),
            "macro_f1": float(f1_score(y_test, predictions, average="macro", zero_division=0)),
            "classification_report": classification_report(
                y_test,
                predictions,
                labels=CLASS_ORDER,
                output_dict=True,
                zero_division=0,
            ),
            "confusion_matrix": confusion_matrix(
                y_test, predictions, labels=CLASS_ORDER
            ).tolist(),
        }
        prediction_rows.extend(
            {
                "customer_id": customer_id,
                "actual_segment": actual,
                "predicted_segment": predicted,
                "model": name,
            }
            for customer_id, actual, predicted in zip(id_test, y_test, predictions)
        )

    selected_name = max(
        metrics,
        key=lambda name: (metrics[name]["macro_f1"], metrics[name]["accuracy"]),
    )
    metrics_dir = ROOT / "outputs" / "metrics"
    metrics_dir.mkdir(parents=True, exist_ok=True)
    results = {
        "synthetic_data": True,
        "target": (
            "A simulated latent engagement propensity, divided at its 33rd and 67th "
            "percentiles into Low, Medium, and High. The latent value is excluded "
            "from all predictor features."
        ),
        "feature_names": list(x.columns),
        "categorical_features": categorical,
        "numeric_features": numeric,
        "training_records": int(len(x_train)),
        "test_records": int(len(x_test)),
        "test_size": 0.25,
        "random_state": 42,
        "class_order": CLASS_ORDER,
        "models": metrics,
        "selected_model": selected_name,
        "selection_reason": (
            "Highest test macro F1-score (tie-breaker: accuracy), prioritizing balanced "
            "performance across the three engagement classes over the majority-class baseline."
        ),
    }
    (metrics_dir / "model_metrics.json").write_text(
        json.dumps(results, indent=2), encoding="utf-8"
    )
    report_text = classification_report(
        y_test,
        fitted[selected_name].predict(x_test),
        labels=CLASS_ORDER,
        target_names=CLASS_ORDER,
        zero_division=0,
    )
    (metrics_dir / "classification_report.txt").write_text(
        f"Synthetic-data test-set report for {selected_name}\n\n{report_text}",
        encoding="utf-8",
    )
    selected_confusion = np.asarray(metrics[selected_name]["confusion_matrix"])
    pd.DataFrame(
        selected_confusion,
        index=[f"actual_{label}" for label in CLASS_ORDER],
        columns=[f"predicted_{label}" for label in CLASS_ORDER],
    ).to_csv(metrics_dir / "confusion_matrix.csv", index=True)
    pd.DataFrame(prediction_rows).to_csv(
        metrics_dir / "test_predictions.csv", index=False
    )
    joblib.dump(fitted[selected_name], metrics_dir / "selected_model.joblib")
    return results


if __name__ == "__main__":
    print(json.dumps(run_round2(), indent=2))
