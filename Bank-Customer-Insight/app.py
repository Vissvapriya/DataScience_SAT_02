"""Interactive Streamlit prototype for the synthetic TrustWave Bank analysis."""

from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parent
PROCESSED = ROOT / "data" / "processed"

st.set_page_config(page_title="TrustWave Bank | Customer Insight", layout="wide")
st.title("TrustWave Bank — Customer Insight")
st.caption("Synthetic-data prototype only. These findings do not describe real bank customers.")

required_paths = [
    PROCESSED / "customer_insight_features.csv",
    PROCESSED / "customer_engagement_target.csv",
    PROCESSED / "account_activity.csv",
    PROCESSED / "channel_analysis.csv",
    PROCESSED / "complaint_analysis.csv",
    PROCESSED / "service_trends.csv",
    ROOT / "outputs" / "metrics" / "round1_findings.json",
]
missing = [str(path.relative_to(ROOT)) for path in required_paths if not path.exists()]
if missing:
    st.error("Run `python main.py` before opening the dashboard. Missing: " + ", ".join(missing))
    st.stop()

features = pd.read_csv(PROCESSED / "customer_insight_features.csv")
target = pd.read_csv(PROCESSED / "customer_engagement_target.csv")
segments = features.merge(target, on="customer_id", validate="one_to_one")
account_activity = pd.read_csv(PROCESSED / "account_activity.csv")
channel_analysis = pd.read_csv(PROCESSED / "channel_analysis.csv")
complaints = pd.read_csv(PROCESSED / "complaint_analysis.csv")
service_trends = pd.read_csv(PROCESSED / "service_trends.csv")
findings = json.loads((ROOT / "outputs" / "metrics" / "round1_findings.json").read_text(encoding="utf-8"))

left, right = st.columns(2)
with left:
    age_options = ["All"] + sorted(segments["age_band"].dropna().unique().tolist())
    selected_age = st.selectbox("Age band", age_options)
with right:
    region_options = ["All"] + sorted(segments["region"].dropna().unique().tolist())
    selected_region = st.selectbox("Region", region_options)

filtered = segments.copy()
if selected_age != "All":
    filtered = filtered.loc[filtered["age_band"] == selected_age]
if selected_region != "All":
    filtered = filtered.loc[filtered["region"] == selected_region]

metric_columns = st.columns(4)
metric_columns[0].metric("Customers in view", f"{len(filtered):,}")
metric_columns[1].metric(
    "High engagement",
    f"{(filtered['engagement_segment'] == 'High').mean():.1%}" if len(filtered) else "0.0%",
)
metric_columns[2].metric(
    "Average monthly transactions",
    f"{filtered['transaction_frequency_per_month'].mean():.2f}" if len(filtered) else "0.00",
)
metric_columns[3].metric(
    "Complaints in view", f"{int(filtered['complaint_count'].sum()):,}"
)

channel_left, channel_right = st.columns(2)
with channel_left:
    st.subheader("Banking channel usage")
    st.bar_chart(channel_analysis.set_index("channel")["uses"])
    st.caption(f"Most used overall: {findings['top_channel'].title()} ({findings['top_channel_uses']:,} uses).")
with channel_right:
    st.subheader("Customer engagement segments")
    segment_counts = filtered["engagement_segment"].value_counts().reindex(["Low", "Medium", "High"], fill_value=0)
    st.bar_chart(segment_counts)

activity_left, activity_right = st.columns(2)
with activity_left:
    st.subheader("Account activity")
    activity_chart = account_activity.set_index("account_type")[["Q1", "Q4"]]
    st.bar_chart(activity_chart)
    st.caption("Quarterly counts compare January–March with October–December 2025.")
with activity_right:
    st.subheader("Complaints by region")
    st.bar_chart(complaints.set_index("region")["complaints"])

st.subheader("Service request trends")
st.bar_chart(service_trends.set_index("request_type")[["requests", "complaints"]])
st.dataframe(service_trends, use_container_width=True, hide_index=True)
