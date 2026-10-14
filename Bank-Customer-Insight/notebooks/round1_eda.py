"""Synthetic source-data generation, cleaning, EDA, and business-question answers."""

from __future__ import annotations

import json
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

ROOT = Path(__file__).resolve().parents[1]
SEED = 20261013
AGE_BANDS = ["18-29", "30-44", "45-59", "60+"]
CHANNELS = ["branch", "atm", "mobile", "internet"]
ACCOUNT_TYPES = ["Checking", "Savings", "Credit Card", "Term Deposit"]
REGIONS = ["North", "South", "East", "West", "Central"]

sns.set_theme(style="whitegrid", palette="Set2")


def _json_value(value: object) -> object:
    """Convert NumPy/Pandas values into JSON-safe Python values."""
    if isinstance(value, (np.integer,)):
        return int(value)
    if isinstance(value, (np.floating,)):
        return float(value)
    if isinstance(value, (np.bool_,)):
        return bool(value)
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    raise TypeError(f"Cannot serialize {type(value).__name__}")


def generate_synthetic_data() -> dict[str, pd.DataFrame]:
    """Generate related, reproducible data and a latent engagement label."""
    rng = np.random.default_rng(SEED)
    n_customers = 1200
    customer_ids = [f"C{i:06d}" for i in range(1, n_customers + 1)]
    age_band = rng.choice(AGE_BANDS, n_customers, p=[0.20, 0.34, 0.28, 0.18])
    region = rng.choice(REGIONS, n_customers, p=[0.22, 0.21, 0.19, 0.20, 0.18])
    relationship_months = rng.integers(1, 301, n_customers)
    latent_engagement = (
        rng.normal(0, 1, n_customers)
        + (relationship_months - relationship_months.mean()) / relationship_months.std() * 0.12
    )
    low_cut, high_cut = np.quantile(latent_engagement, [1 / 3, 2 / 3])
    target = np.where(
        latent_engagement < low_cut,
        "Low",
        np.where(latent_engagement < high_cut, "Medium", "High"),
    )

    customers = pd.DataFrame(
        {
            "customer_id": customer_ids,
            "age_band": age_band,
            "region": region,
            "relationship_months": relationship_months,
        }
    )

    account_rows: list[dict[str, object]] = []
    accounts_by_customer: dict[str, list[tuple[str, str]]] = {}
    account_number = 1
    for customer_id in customer_ids:
        count = int(rng.choice([1, 2, 3], p=[0.48, 0.44, 0.08]))
        account_types = rng.choice(ACCOUNT_TYPES, count, replace=False)
        accounts_by_customer[customer_id] = []
        customer_index = int(customer_id[1:]) - 1
        for account_type in account_types:
            account_id = f"A{account_number:07d}"
            account_number += 1
            base_balance = {
                "Checking": 2400,
                "Savings": 9800,
                "Credit Card": 1250,
                "Term Deposit": 18400,
            }[account_type]
            balance = rng.lognormal(
                np.log(base_balance) + latent_engagement[customer_index] * 0.12, 0.72
            )
            account_rows.append(
                {
                    "account_id": account_id,
                    "customer_id": customer_id,
                    "account_type": account_type,
                    "balance": round(float(balance), 2),
                }
            )
            accounts_by_customer[customer_id].append((account_id, account_type))
    accounts = pd.DataFrame(account_rows)

    transactions: list[dict[str, object]] = []
    transaction_number = 1
    transaction_months = pd.date_range("2025-01-01", periods=12, freq="MS")
    for customer_index, customer_id in enumerate(customer_ids):
        customer_accounts = accounts_by_customer[customer_id]
        for month_index, month in enumerate(transaction_months):
            monthly_mean = np.exp(1.16 + 0.25 * latent_engagement[customer_index])
            transaction_count = int(rng.poisson(monthly_mean))
            progress = month_index / 11
            for _ in range(transaction_count):
                weighted_accounts = []
                for _, account_type in customer_accounts:
                    trend_factor = {
                        "Checking": 1.0,
                        "Savings": 1.0,
                        "Credit Card": 0.72 + 0.70 * progress,
                        "Term Deposit": 1.38 - 0.78 * progress,
                    }[account_type]
                    weighted_accounts.append(trend_factor)
                selected_index = int(rng.choice(len(customer_accounts), p=np.array(weighted_accounts) / sum(weighted_accounts)))
                account_id, account_type = customer_accounts[selected_index]
                start = month
                end = month + pd.offsets.MonthBegin(1)
                days_in_month = (end - start).days
                transaction_date = start + pd.Timedelta(
                    days=int(rng.integers(0, days_in_month)),
                    seconds=int(rng.integers(0, 86400)),
                )
                amount = float(rng.lognormal(4.25, 0.95))
                if rng.random() < 0.002:
                    amount *= float(rng.uniform(35, 110))
                transaction_type = rng.choice(
                    ["Payment", "Withdrawal", "Transfer", "Deposit", "Purchase"],
                    p=[0.22, 0.16, 0.21, 0.17, 0.24],
                )
                transactions.append(
                    {
                        "transaction_id": f"T{transaction_number:09d}",
                        "transaction_date": transaction_date,
                        "amount": round(amount, 2),
                        "transaction_type": transaction_type,
                        "account_id": account_id,
                    }
                )
                transaction_number += 1
    transactions_df = pd.DataFrame(transactions)

    inconsistent_labels = {
        "Checking": ["Checking", "CHECKING", "Current", "Checking Account"],
        "Savings": ["Savings", "SAVINGS ", "Savings Account"],
        "Credit Card": ["Credit Card", "CREDIT CARD", "Credit"],
        "Term Deposit": ["Term Deposit", "Fixed Deposit", "TERM DEPOSIT "],
    }
    accounts["account_type"] = [
        str(rng.choice(inconsistent_labels[value]))
        for value in accounts["account_type"]
    ]

    channel_rows: list[dict[str, object]] = []
    for index, customer_id in enumerate(customer_ids):
        age = age_band[index]
        digital_factor = {"18-29": 1.55, "30-44": 1.25, "45-59": 0.86, "60+": 0.55}[age]
        propensity = float(latent_engagement[index])
        row: dict[str, object] = {"customer_id": customer_id}
        means = {
            "branch": max(0.15, 2.8 - 0.22 * propensity - digital_factor * 0.25),
            "atm": max(0.15, 4.1 - 0.15 * propensity - digital_factor * 0.45),
            "mobile": max(0.15, 4.0 * digital_factor * np.exp(0.18 * propensity)),
            "internet": max(0.15, 2.5 * digital_factor * np.exp(0.15 * propensity)),
        }
        for channel, mean in means.items():
            value: object = int(rng.poisson(mean))
            if rng.random() < 0.055:
                value = np.nan
            row[f"{channel}_uses"] = value
        channel_rows.append(row)
    channels = pd.DataFrame(channel_rows)

    service_rows: list[dict[str, object]] = []
    service_number = 1
    service_categories = ["Card support", "Account inquiry", "Digital access", "Loan service", "Payments"]
    for index, customer_id in enumerate(customer_ids):
        request_count = int(rng.poisson(1.65 + 0.10 * max(-1.5, -latent_engagement[index])))
        for _ in range(request_count):
            is_complaint = bool(
                rng.random()
                < min(0.72, 0.27 + (0.11 if region[index] in {"North", "West"} else 0) - 0.04 * latent_engagement[index])
            )
            row = {
                "service_id": f"S{service_number:08d}",
                "customer_id": customer_id,
                "request_date": pd.Timestamp("2025-01-01")
                + pd.Timedelta(days=int(rng.integers(0, 365))),
                "request_type": str(rng.choice(service_categories)),
                "is_complaint": is_complaint,
                "resolution_status": str(rng.choice(["Resolved", "Pending", "Escalated"], p=[0.78, 0.13, 0.09])),
                "resolution_days": round(float(rng.gamma(2.0, 1.8)), 1),
            }
            if is_complaint and rng.random() < 0.11:
                row["request_type"] = np.nan
            if rng.random() < 0.07:
                row["resolution_status"] = np.nan
            if rng.random() < 0.09:
                row["resolution_days"] = np.nan
            service_rows.append(row)
            service_number += 1
    service = pd.DataFrame(service_rows)

    if not transactions_df.empty:
        duplicate_count = max(1, int(len(transactions_df) * 0.01))
        transactions_df = pd.concat(
            [transactions_df, transactions_df.sample(duplicate_count, random_state=SEED)],
            ignore_index=True,
        )

    return {
        "customers": customers,
        "accounts": accounts,
        "transactions": transactions_df,
        "channels": channels,
        "service": service,
        "target": pd.DataFrame({"customer_id": customer_ids, "engagement_segment": target}),
    }


def _quality_summary(frame: pd.DataFrame) -> dict[str, object]:
    return {
        "records": int(len(frame)),
        "attributes": list(frame.columns),
        "attribute_count": int(frame.shape[1]),
        "data_types": {name: str(dtype) for name, dtype in frame.dtypes.items()},
        "missing_values": {name: int(value) for name, value in frame.isna().sum().items()},
        "missing_cells": int(frame.isna().sum().sum()),
        "duplicate_rows": int(frame.duplicated().sum()),
    }


def _standardize_account_label(value: object) -> str:
    normalized = re.sub(r"[^a-z]", "", str(value).lower())
    if normalized in {"checking", "checkingaccount", "current", "currentaccount", "checkings"}:
        return "Checking"
    if normalized in {"savings", "savingsaccount", "saving"}:
        return "Savings"
    if normalized in {"credit", "creditcard", "creditcards"}:
        return "Credit Card"
    if normalized in {"termdeposit", "fixeddeposit", "deposit"}:
        return "Term Deposit"
    raise ValueError(f"Unrecognized account type: {value!r}")


def _save_chart(fig: plt.Figure, filename: str) -> None:
    chart_dir = ROOT / "outputs" / "charts"
    chart_dir.mkdir(parents=True, exist_ok=True)
    fig.tight_layout()
    fig.savefig(chart_dir / filename, dpi=150, bbox_inches="tight")
    plt.close(fig)


def run_round1() -> dict[str, object]:
    """Generate, clean, report on, and analyze the five source datasets."""
    raw_dir = ROOT / "data" / "raw"
    processed_dir = ROOT / "data" / "processed"
    raw_dir.mkdir(parents=True, exist_ok=True)
    processed_dir.mkdir(parents=True, exist_ok=True)
    frames = generate_synthetic_data()

    raw_names = ["customers", "accounts", "transactions", "channels", "service"]
    before = {}
    for name in raw_names:
        path = raw_dir / f"{name}.csv"
        frames[name].to_csv(path, index=False)
        before[name] = _quality_summary(pd.read_csv(path))
    target = frames["target"]

    customers = frames["customers"].copy()
    accounts = frames["accounts"].copy()
    accounts["account_type"] = accounts["account_type"].map(_standardize_account_label)
    transactions = frames["transactions"].copy()
    transaction_rows_before = len(transactions)
    transactions = transactions.drop_duplicates(subset=["transaction_id"], keep="first").copy()
    transactions["transaction_date"] = pd.to_datetime(transactions["transaction_date"])
    transactions["amount_original"] = pd.to_numeric(transactions["amount"], errors="raise")
    q1, q3 = transactions["amount_original"].quantile([0.25, 0.75])
    iqr = q3 - q1
    lower_bound, upper_bound = max(0.0, q1 - 1.5 * iqr), q3 + 1.5 * iqr
    transactions["is_amount_outlier"] = ~transactions["amount_original"].between(lower_bound, upper_bound)
    outlier_count = int(transactions["is_amount_outlier"].sum())
    transactions["amount"] = transactions["amount_original"].clip(lower_bound, upper_bound).round(2)

    channels = frames["channels"].copy()
    channels = channels.merge(customers[["customer_id", "age_band"]], on="customer_id", validate="one_to_one")
    for column in [f"{channel}_uses" for channel in CHANNELS]:
        age_median = channels.groupby("age_band")[column].transform("median")
        channels[column] = channels[column].fillna(age_median).fillna(channels[column].median()).round().astype(int)
    channels = channels.drop(columns="age_band")

    service = frames["service"].copy()
    service["request_date"] = pd.to_datetime(service["request_date"])
    service["request_type"] = service["request_type"].fillna("Unspecified")
    service["resolution_status"] = service["resolution_status"].fillna("Unknown")
    resolution_median = service["resolution_days"].median()
    service["resolution_days"] = service["resolution_days"].fillna(resolution_median).round(1)
    service["is_complaint"] = service["is_complaint"].astype(bool)

    cleaned = {
        "customers": customers,
        "accounts": accounts,
        "transactions": transactions,
        "channels": channels,
        "service": service,
    }
    after = {}
    for name, frame in cleaned.items():
        frame.to_csv(processed_dir / f"{name}_clean.csv", index=False)
        after[name] = _quality_summary(frame)
    target.to_csv(processed_dir / "customer_engagement_target.csv", index=False)

    clean_account_transactions = transactions.merge(
        accounts[["account_id", "customer_id", "account_type"]],
        on="account_id",
        how="left",
        validate="many_to_one",
    )
    if clean_account_transactions["customer_id"].isna().any():
        raise ValueError("Some transactions reference account IDs absent from accounts.csv.")
    clean_service = service.merge(
        customers[["customer_id", "region", "age_band"]],
        on="customer_id",
        how="left",
        validate="many_to_one",
    )
    account_month = clean_account_transactions.assign(
        month=clean_account_transactions["transaction_date"].dt.to_period("M").dt.to_timestamp()
    )

    channel_totals = channels[[f"{channel}_uses" for channel in CHANNELS]].sum().sort_values(ascending=False)
    channel_by_age = channels.merge(
        customers[["customer_id", "age_band"]], on="customer_id", validate="one_to_one"
    ).groupby("age_band")[[f"{channel}_uses" for channel in CHANNELS]].mean()
    account_quarter = clean_account_transactions.assign(
        quarter=np.select(
            [
                clean_account_transactions["transaction_date"].dt.month <= 3,
                clean_account_transactions["transaction_date"].dt.month >= 10,
            ],
            ["Q1", "Q4"],
            default="Other",
        )
    )
    account_quarter = account_quarter[account_quarter["quarter"].isin(["Q1", "Q4"])]
    account_activity = (
        account_quarter.groupby(["account_type", "quarter"])
        .size()
        .unstack(fill_value=0)
        .reindex(columns=["Q1", "Q4"], fill_value=0)
    )
    account_activity["q1_to_q4_change_pct"] = np.where(
        account_activity["Q1"] > 0,
        (account_activity["Q4"] - account_activity["Q1"]) / account_activity["Q1"] * 100,
        np.nan,
    )
    regional_complaints = (
        clean_service.loc[clean_service["is_complaint"]]
        .groupby("region")
        .size()
        .sort_values(ascending=False)
    )
    customer_balance = accounts.groupby("customer_id")["balance"].sum().rename("total_balance")
    balance_by_segment = (
        customer_balance.to_frame()
        .join(target.set_index("customer_id"))
        .groupby("engagement_segment")["total_balance"]
        .agg(["mean", "median", "count"])
        .reindex(["Low", "Medium", "High"])
    )
    repeated_by_type = (
        service.groupby(["customer_id", "request_type"])
        .size()
        .loc[lambda counts: counts > 1]
        .groupby("request_type")
        .size()
        .sort_values(ascending=False)
    )
    monthly_transactions = account_month.groupby(account_month["transaction_date"].dt.to_period("M").astype(str)).size()

    findings = {
        "synthetic_data": True,
        "seed": SEED,
        "source_record_counts": {name: len(frames[name]) for name in raw_names},
        "duplicate_transaction_rows_removed": int(transaction_rows_before - len(transactions)),
        "transaction_amount_iqr_bounds": {"lower": float(lower_bound), "upper": float(upper_bound)},
        "transaction_amount_outliers_flagged_and_capped": outlier_count,
        "missing_channel_values_filled": int(before["channels"]["missing_cells"]),
        "channel_totals": {channel.replace("_uses", ""): int(value) for channel, value in channel_totals.items()},
        "top_channel": str(channel_totals.index[0]).replace("_uses", ""),
        "top_channel_uses": int(channel_totals.iloc[0]),
        "channel_average_uses_by_age_band": channel_by_age.round(2).to_dict(orient="index"),
        "account_activity_q1_q4": {
            account_type: {
                "q1_transactions": int(row["Q1"]),
                "q4_transactions": int(row["Q4"]),
                "change_percent": round(float(row["q1_to_q4_change_pct"]), 2),
            }
            for account_type, row in account_activity.iterrows()
        },
        "declining_account_types": [
            str(account_type)
            for account_type, row in account_activity.iterrows()
            if row["q1_to_q4_change_pct"] < 0
        ],
        "regional_complaint_counts": {str(k): int(v) for k, v in regional_complaints.items()},
        "highest_complaint_region": str(regional_complaints.index[0]) if len(regional_complaints) else "None",
        "average_balance_by_engagement_segment": {
            str(segment): round(float(row["mean"]), 2)
            for segment, row in balance_by_segment.iterrows()
        },
        "repeated_requests_by_service_type": {str(k): int(v) for k, v in repeated_by_type.items()},
        "most_repeated_service_type": str(repeated_by_type.index[0]) if len(repeated_by_type) else "None",
        "monthly_transaction_counts": {str(k): int(v) for k, v in monthly_transactions.items()},
        "data_quality": {"before": before, "after": after},
        "synthetic_target_definition": (
            "Low/Medium/High are tertiles of a simulated latent engagement propensity. "
            "The latent propensity influences transaction and digital-channel behavior, "
            "but is not included as a model feature."
        ),
    }
    metrics_dir = ROOT / "outputs" / "metrics"
    metrics_dir.mkdir(parents=True, exist_ok=True)
    (metrics_dir / "round1_findings.json").write_text(
        json.dumps(findings, indent=2, default=_json_value), encoding="utf-8"
    )
    (metrics_dir / "data_quality_report.json").write_text(
        json.dumps({"synthetic_data": True, **findings["data_quality"]}, indent=2), encoding="utf-8"
    )
    (processed_dir / "account_activity.csv").write_text(
        account_activity.reset_index().to_csv(index=False), encoding="utf-8"
    )
    (processed_dir / "channel_analysis.csv").write_text(
        channel_totals.rename_axis("channel").reset_index(name="uses").to_csv(index=False), encoding="utf-8"
    )
    (processed_dir / "balance_analysis.csv").write_text(
        balance_by_segment.reset_index().to_csv(index=False), encoding="utf-8"
    )
    (processed_dir / "complaint_analysis.csv").write_text(
        regional_complaints.rename_axis("region").reset_index(name="complaints").to_csv(index=False), encoding="utf-8"
    )
    service_trends = (
        service.groupby("request_type")
        .agg(requests=("service_id", "count"), complaints=("is_complaint", "sum"))
        .sort_values("requests", ascending=False)
        .reset_index()
    )
    service_trends.to_csv(processed_dir / "service_trends.csv", index=False)

    fig, ax = plt.subplots(figsize=(8, 5))
    channel_totals.rename(index=lambda value: value.replace("_uses", "").title()).plot.bar(ax=ax)
    ax.set(title="Synthetic Banking Channel Usage", xlabel="Channel", ylabel="Total reported uses")
    _save_chart(fig, "channel_usage.png")

    fig, ax = plt.subplots(figsize=(9, 5))
    channel_by_age.rename(columns=lambda value: value.replace("_uses", "").title()).plot.bar(ax=ax)
    ax.set(title="Average Channel Use by Age Band", xlabel="Age band", ylabel="Average uses per customer")
    ax.legend(title="Channel", bbox_to_anchor=(1.02, 1), loc="upper left")
    _save_chart(fig, "age_group_comparison.png")

    fig, ax = plt.subplots(figsize=(9, 5))
    account_activity[["Q1", "Q4"]].plot.bar(ax=ax)
    ax.set(title="Account Activity: Q1 vs Q4 2025", xlabel="Account type", ylabel="Transactions")
    _save_chart(fig, "account_activity.png")

    fig, ax = plt.subplots(figsize=(8, 5))
    regional_complaints.sort_values().plot.barh(ax=ax)
    ax.set(title="Complaints by Region", xlabel="Complaint records", ylabel="Region")
    _save_chart(fig, "regional_complaints.png")

    plot_balances = customer_balance.to_frame().join(target.set_index("customer_id"))
    fig, ax = plt.subplots(figsize=(8, 5))
    sns.boxplot(data=plot_balances, x="engagement_segment", y="total_balance", order=["Low", "Medium", "High"], ax=ax)
    ax.set(title="Customer Balance by Synthetic Engagement Segment", xlabel="Segment", ylabel="Total balance")
    _save_chart(fig, "balance_distribution.png")

    fig, ax = plt.subplots(figsize=(10, 5))
    monthly_transactions.plot(ax=ax, marker="o")
    ax.set(title="Monthly Transaction Trend (2025)", xlabel="Month", ylabel="Transaction records")
    ax.tick_params(axis="x", rotation=45)
    _save_chart(fig, "transaction_trends.png")

    fig, ax = plt.subplots(figsize=(9, 5))
    service_trends.set_index("request_type")["requests"].sort_values().plot.barh(ax=ax)
    ax.set(title="Service Requests by Type", xlabel="Request records", ylabel="Service type")
    _save_chart(fig, "service_requests.png")

    return findings


if __name__ == "__main__":
    result = run_round1()
    print(json.dumps(result, indent=2, default=_json_value))
