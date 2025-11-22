import streamlit as st
import pandas as pd
import numpy as np

# =========================
# CONFIG
# =========================
EXCEL_PATH = "dynamic rebate system.xlsx"  # adjust if needed

st.set_page_config(
    page_title="Dynamic Rebate Dashboard",
    layout="wide",
)

# =========================
# LOAD & CLEAN DATA
# =========================
@st.cache_data
def load_sheet1(path: str) -> pd.DataFrame:
    """
    Task-level sheet.
    Read everything as string (dtype=str) to avoid pyarrow ArrowTypeError,
    then manually convert numeric columns.
    """
    df = pd.read_excel(path, sheet_name="Sheet1", dtype=str)

    # Helper: clean numeric-looking strings into float
    def clean_numeric(series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.replace(r"[^0-9.\-]", "", regex=True)
            .replace({"": np.nan, "nan": np.nan, "-": np.nan})
            .astype(float)
        )

    numeric_cols = [
        "Deposited Amount (USD)",
        "Required Deposit Amount (USD)",
        "Completed Lots",
        "Required number of lots completed",
        "Number of completed registrations",
        "Required number of completed registrations",
        "Number of completed KYC",
        "Required number of completed KYC",
        "Number of completed activations",
        "Required number of completed activations",
        "Task amount (USD)",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col + "_num"] = clean_numeric(df[col])

    # Period as string
    if "Period" in df.columns:
        df["Period"] = df["Period"].astype(str)

    # Completed flag
    if "Task Status" in df.columns:
        df["is_completed"] = (df["Task Status"] == "Completed").astype(int)
    else:
        df["is_completed"] = 0

    return df


@st.cache_data
def load_sheet2(path: str) -> pd.DataFrame:
    """
    Pool / POG sheet.
    Read as string then convert to numeric for key columns.
    """
    df = pd.read_excel(path, sheet_name="Sheet2", dtype=str)

    df = df.rename(
        columns={
            "Period": "Period",
            "HPOV PL": "HPOV_PL",
            "POG": "POG",
            "Used Pool": "Used_Pool",
            "Given / Pool %": "Given_Pool_Ratio",
        }
    )

    # Convert numeric columns safely
    for col in ["HPOV_PL", "POG", "Used_Pool", "Given_Pool_Ratio"]:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", ""), errors="coerce"
            )

    if "Period" in df.columns:
        df["Period"] = df["Period"].astype(str)

    return df


@st.cache_data
def load_sheet3(path: str) -> pd.DataFrame:
    """
    Campaign summary sheet (aggregated per period).
    Skip first 2 header rows, then clean/rename.
    Everything read as string, then convert numerics manually.
    """
    raw = pd.read_excel(path, sheet_name="Sheet3", skiprows=2, dtype=str)

    df = raw.rename(
        columns={
            "Period": "Period",
            "Unnamed: 1": "num_users",
            "Unnamed: 2": "num_users_completed_all",
            "Unnamed: 3": "tv_total_users",
            "Unnamed: 4": "tv_users_completed",
            "Unnamed: 5": "tv_required_lots",
            "TV lot": "tv_all_status_lots",
            "Unnamed: 7": "tv_completed_lots",
            "TV task bonus $": "tv_bonus_str",
            "Unnamed: 9": "dp_users_completed",
            "Unnamed: 10": "dp_required_deposit",
            "Unnamed: 11": "dp_total_deposit_str",
            "Unnamed: 12": "dp_completed_deposit_str",
            "Unnamed: 13": "dp_bonus_str",
            "DP task bonus $": "dp_bonus2_str",
            "Unnamed: 15": "ref_finished_users",
            "Unnamed: 16": "ref_completed_reg",
            "Unnamed: 17": "ref_total_reg",
            "Unnamed: 18": "ref_completed_kyc",
            "Unnamed: 19": "ref_total_kyc",
            "Unnamed: 20": "ref_completed_activation",
            "Unnamed: 21": "ref_total_activation",
            "Unnamed: 22": "ref_bonus_str",
            "Unnamed: 23": "bb_users_completed",
            "Rf task bonus $": "bb_bonus_str",
            "Blind Box": "bb_bonus2_str",
            "Farrah Bonus": "farrah_bonus_str",
            "Total Bonus": "total_bonus_str",
            "Actual Bonus": "actual_bonus_str",
        }
    )

    if "Period" in df.columns:
        df["Period"] = df["Period"].astype(str)

    # Money cleaner
    def clean_money(series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.replace(r"[^0-9.\-]", "", regex=True)
            .replace({"": np.nan, "nan": np.nan})
            .astype(float)
        )

    money_cols = [
        "tv_bonus_str",
        "dp_total_deposit_str",
        "dp_completed_deposit_str",
        "dp_bonus_str",
        "dp_bonus2_str",
        "ref_bonus_str",
        "bb_bonus_str",
        "bb_bonus2_str",
        "farrah_bonus_str",
        "total_bonus_str",
        "actual_bonus_str",
    ]
    for col in money_cols:
        if col in df.columns:
            df[col.replace("_str", "_num")] = clean_money(df[col])

    # Convert numeric columns (counts / lots / etc.)
    numeric_cols = [
        "num_users",
        "num_users_completed_all",
        "tv_total_users",
        "tv_users_completed",
        "tv_required_lots",
        "tv_all_status_lots",
        "tv_completed_lots",
        "dp_users_completed",
        "dp_required_deposit",
        "ref_finished_users",
        "ref_completed_reg",
        "ref_total_reg",
        "ref_completed_kyc",
        "ref_total_kyc",
        "ref_completed_activation",
        "ref_total_activation",
        "bb_users_completed",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(",", ""), errors="coerce"
            )

    return df


# Load once
sheet1 = load_sheet1(EXCEL_PATH)
sheet2 = load_sheet2(EXCEL_PATH)
sheet3 = load_sheet3(EXCEL_PATH)


# =========================
# STREAMLIT LAYOUT
# =========================


st.title("Dynamic Rebate System Dashboard")

tab1, tab2, tab3 = st.tabs(
    ["Task Level", "Pool & POG", "Campaign Summary"]
)

# ======================================================
# TAB 1 – SHEET1: TASK-LEVEL ANALYSIS (NO RAW DATA)
# ======================================================
with tab1:
    st.header("Task Level Analysis")

    # Filters
    period_options = sorted(sheet1["Period"].unique())
    task_types = sorted(sheet1["Task Type"].unique())
    statuses = sorted(sheet1["Task Status"].unique())

    c1, c2, c3 = st.columns(3)
    with c1:
        sel_periods = st.multiselect(
            "Period(s)", period_options, default=period_options
        )
    with c2:
        sel_types = st.multiselect("Task Type", task_types, default=task_types)
    with c3:
        sel_status = st.multiselect("Status", statuses, default=statuses)

    filtered1 = sheet1[
        sheet1["Period"].isin(sel_periods)
        & sheet1["Task Type"].isin(sel_types)
        & sheet1["Task Status"].isin(sel_status)
    ].copy()

    # KPIs
    total_tasks = len(filtered1)
    completion_rate = (
        filtered1["is_completed"].mean() if total_tasks > 0 else 0.0
    )
    total_task_amount = filtered1["Task amount (USD)_num"].sum()

    k1, k2, k3 = st.columns(3)
    k1.metric("Total Tasks", f"{total_tasks:,}")
    k2.metric("Completion Rate", f"{completion_rate*100:.2f}%")
    

    # Completion by Task Type
    type_summary = (
        filtered1.groupby("Task Type")
        .agg(total=("Client ID", "count"), completed=("is_completed", "sum"))
        .assign(rate=lambda x: x["completed"] / x["total"])
    )

    if not type_summary.empty:
        st.markdown("### Completion Rate by Task Type")
        st.bar_chart(type_summary["rate"] * 100)
    else:
        st.info("No data for current filter combination.")

    # Tasks & completion by Period
    period_summary = (
        filtered1.groupby("Period")
        .agg(total=("Client ID", "count"), completed=("is_completed", "sum"))
        .assign(rate=lambda x: x["completed"] / x["total"])
        .sort_values("Period")
    )

    if not period_summary.empty:
        c1, c2 = st.columns(2)
        with c1:
            st.write("Tasks per Period")
            st.bar_chart(period_summary["total"])
        with c2:
            st.write("Completion % per Period")
            st.line_chart(period_summary["rate"] * 100)
    else:
        st.info("No period-level data for current filters.")


# ======================================================
# TAB 2 – SHEET2: POOL & POG (NO RAW DATA)
# ======================================================
with tab2:
    st.header("Pool & POG Overview")

    total_pog = sheet2["POG"].sum()
    total_used_pool = sheet2["Used_Pool"].sum()
    avg_ratio = sheet2["Given_Pool_Ratio"].mean()

    k1, k2, k3 = st.columns(3)
    k1.metric("Total POG", f"{total_pog:,.2f}")
    k2.metric("Total Used Pool", f"{total_used_pool:,.2f}")
    k3.metric(
        "Avg Given / Pool %",
        f"{avg_ratio:.2f}%" if not np.isnan(avg_ratio) else "N/A",
    )

    sheet2_sorted = sheet2.sort_values("Period")

    c1, c2 = st.columns(2)
    with c1:
        st.write("HPOV & POG per Period")
        st.line_chart(
            sheet2_sorted.set_index("Period")[["HPOV_PL", "POG"]]
        )
    with c2:
        st.write("Used Pool per Period")
        st.bar_chart(sheet2_sorted.set_index("Period")["Used_Pool"])

    st.write("Given/Pool % per Period")
    st.bar_chart(
        sheet2_sorted.set_index("Period")["Given_Pool_Ratio"] 
    )


# ======================================================
# TAB 3 – SHEET3: CAMPAIGN SUMMARY (NO RAW DATA)
# ======================================================
with tab3:
    st.header("Campaign Summary")

    df3 = sheet3.copy().sort_values("Period")

    # Derived metrics
    if {"num_users", "num_users_completed_all"}.issubset(df3.columns):
        df3["user_completion_rate"] = (
            df3["num_users_completed_all"] / df3["num_users"]
        )
    else:
        df3["user_completion_rate"] = np.nan

    total_users = df3.get("num_users", pd.Series(dtype=float)).sum()
    total_completed = df3.get(
        "num_users_completed_all", pd.Series(dtype=float)
    ).sum()
    avg_rate = (
        df3["user_completion_rate"].mean()
        if "user_completion_rate" in df3.columns
        else np.nan
    )
    total_actual_bonus = df3.get(
        "actual_bonus_num", pd.Series(dtype=float)
    ).sum()

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total Users", f"{total_users:,.0f}")
    k2.metric("Users Completed All", f"{total_completed:,.0f}")
    k3.metric(
        "Avg Completion %",
        f"{avg_rate*100:.2f}%" if not np.isnan(avg_rate) else "N/A",
    )
    k4.metric("Total Actual Bonus", f"{total_actual_bonus:,.2f}")

    # User completion visuals
    if {"num_users", "num_users_completed_all"}.issubset(df3.columns):
        st.subheader("User Completion Over Time")
        st.bar_chart(
            df3.set_index("Period")[["num_users", "num_users_completed_all"]]
        )

        if "user_completion_rate" in df3.columns:
            st.line_chart(
                df3.set_index("Period")["user_completion_rate"] * 100
            )

    # Trading volume
    st.subheader("Trading Volume – Lots & Bonus")
    tv_cols = [
        col
        for col in ["tv_all_status_lots", "tv_completed_lots"]
        if col in df3.columns
    ]
    if tv_cols:
        st.line_chart(df3.set_index("Period")[tv_cols])

    if "tv_bonus_num" in df3.columns:
        st.bar_chart(df3.set_index("Period")["tv_bonus_num"])

    # Bonus totals
    bonus_cols = []
    if "total_bonus_num" in df3.columns:
        bonus_cols.append("total_bonus_num")
    if "actual_bonus_num" in df3.columns:
        bonus_cols.append("actual_bonus_num")

    if bonus_cols:
        st.subheader("Bonus Distribution Over Time")
        st.bar_chart(df3.set_index("Period")[bonus_cols])
