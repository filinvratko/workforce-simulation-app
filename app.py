# app.py
import os
import json
import re
from typing import Dict, List, Any, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# App setup
# ============================================================

st.set_page_config(
    page_title="Workforce Planning Simulation Report",
    layout="wide",
    initial_sidebar_state="collapsed",
)

REPORT_TITLE = "Workforce Planning Simulation Report"

DATA_SAMPLE_SHEET = "FTE & Costs"
DRIVERS_SHEET = "SAC Driver"
CALC_RULES_SHEET = "NDC Per country Rules"

DATA_SAMPLE_ROW_COLUMNS = [
    "Company Code",
    "Cost Center",
    "Profit Center",
    "Business Area",
    "Segment",
    "Employee",
    "Account",
]

CALC_RULES_REQUIRED_COLUMNS = [
    "Description",
    "Global Account",
    "SAC Driver",
    "NDC Comment",
    "Israel",
]


# ============================================================
# Mock data
# ============================================================

def generate_mock_source_data() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=24, freq="MS")
    accounts = ["FTE", "Salary USD", "Benefits USD", "Payroll Tax USD"]
    employees = ["E001", "E002", "E003", "E004", "E005"]

    rows = []

    for employee in employees:
        for account in accounts:
            row = {
                "Company Code": "IL01",
                "Cost Center": "CC100",
                "Profit Center": "PC01",
                "Business Area": "Operations",
                "Segment": "Manufacturing",
                "Employee": employee,
                "Account": account,
                "Version": "Baseline",
            }

            for i, month in enumerate(months):
                col = month.strftime("%Y%m")

                if account == "FTE":
                    row[col] = 1.0
                elif account == "Salary USD":
                    row[col] = 5200 + i * 50
                elif account == "Benefits USD":
                    row[col] = 850
                else:
                    row[col] = 620

            rows.append(row)

    return pd.DataFrame(rows)


def generate_mock_drivers() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=24, freq="MS")
    data = {
        "Driver": [
            "Salary Increase",
            "Benefits Increase",
            "Payroll Tax",
            "Hiring Plan",
            "Layoff Plan",
        ]
    }

    for month in months:
        data[month.strftime("%Y%m")] = [0.03, 0.02, 0.12, 0, 0]

    return pd.DataFrame(data)


def generate_mock_rules() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Description": [
                "FTE baseline account",
                "Salary cost calculated from salary driver",
                "Benefits cost calculated from benefits driver",
                "Tax cost calculated from payroll tax driver",
            ],
            "Global Account": ["FTE", "Salary USD", "Benefits USD", "Payroll Tax USD"],
            "SAC Driver": [
                "Hiring Plan",
                "Salary Increase",
                "Benefits Increase",
                "Payroll Tax",
            ],
            "NDC Comment": [
                "Baseline FTE volume",
                "USD salary cost",
                "USD benefits cost",
                "USD tax cost",
            ],
            "Israel": ["Yes", "Yes", "Yes", "Yes"],
        }
    )


# ============================================================
# Excel helpers
# ============================================================

def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        raise ImportError("Excel upload requires openpyxl in requirements.txt.")

    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

    if sheet_name not in sheets:
        raise ValueError(f"Missing required sheet: {sheet_name}")

    return sheets[sheet_name].copy()


def clean_column_name(col: Any) -> str:
    text = str(col).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def is_yyyymm_column(col: Any) -> bool:
    return bool(re.fullmatch(r"\d{6}", clean_column_name(col)))


def normalize_yyyymm_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for col in df.columns:
        clean = clean_column_name(col)
        if re.fullmatch(r"\d{6}", clean):
            rename_map[col] = clean
    return df.rename(columns=rename_map)


def validate_required_columns(
    df: pd.DataFrame,
    required: List[str],
    sheet_name: str,
) -> None:
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}")


def get_month_columns(df: pd.DataFrame, excluded: List[str]) -> List[str]:
    month_cols = [col for col in df.columns if col not in excluded and is_yyyymm_column(col)]
    return sorted(month_cols, key=lambda x: pd.to_datetime(str(x), format="%Y%m"))


def validate_month_columns(df: pd.DataFrame, excluded: List[str], sheet_name: str) -> List[str]:
    month_cols = get_month_columns(df, excluded)

    if not month_cols:
        raise ValueError(
            f"Sheet '{sheet_name}' must contain monthly columns in YYYYMM format, e.g. 202601."
        )

    return month_cols


# ============================================================
# Loaders
# ============================================================

def load_data_sample(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, DATA_SAMPLE_SHEET)
    df = normalize_yyyymm_columns(df)

    validate_required_columns(df, DATA_SAMPLE_ROW_COLUMNS, DATA_SAMPLE_SHEET)

    if "Version" not in df.columns:
        df["Version"] = "Baseline"

    month_cols = validate_month_columns(
        df,
        DATA_SAMPLE_ROW_COLUMNS + ["Version"],
        DATA_SAMPLE_SHEET,
    )

    for col in month_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_drivers(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, DRIVERS_SHEET)
    df = normalize_yyyymm_columns(df)

    validate_required_columns(df, ["Driver"], DRIVERS_SHEET)

    month_cols = validate_month_columns(df, ["Driver"], DRIVERS_SHEET)

    for col in month_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_calculation_rules(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, CALC_RULES_SHEET)
    validate_required_columns(df, CALC_RULES_REQUIRED_COLUMNS, CALC_RULES_SHEET)
    return df


# ============================================================
# Baseline and simulation preparation
# ============================================================

def filter_version(df: pd.DataFrame, version: str) -> pd.DataFrame:
    if "Version" not in df.columns:
        return df.copy()

    version_mask = df["Version"].astype(str).str.lower() == version.lower()
    filtered = df[version_mask].copy()

    if filtered.empty and version.lower() == "baseline":
        return df.copy()

    return filtered


def identify_fte_rows(df: pd.DataFrame) -> pd.Series:
    return df["Account"].astype(str).str.lower().str.contains("fte", na=False)


def identify_usd_cost_rows(df: pd.DataFrame) -> pd.Series:
    account_text = df["Account"].astype(str).str.lower()

    is_fte = account_text.str.contains("fte", na=False)
    is_usd = account_text.str.contains("usd", na=False)

    return (~is_fte) & is_usd


def source_to_monthly_baseline(source_df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_yyyymm_columns(source_df)
    baseline_df = filter_version(df, "Baseline")

    month_cols = get_month_columns(
        baseline_df,
        DATA_SAMPLE_ROW_COLUMNS + ["Version"],
    )

    long_df = baseline_df.melt(
        id_vars=DATA_SAMPLE_ROW_COLUMNS + ["Version"],
        value_vars=month_cols,
        var_name="MonthYYYYMM",
        value_name="Value",
    )

    long_df["Month"] = pd.to_datetime(long_df["MonthYYYYMM"], format="%Y%m")
    long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0)

    fte_df = (
        long_df[identify_fte_rows(long_df)]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "BaselineFTE"})
    )

    usd_cost_df = (
        long_df[identify_usd_cost_rows(long_df)]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "BaselineUSD"})
    )

    monthly = pd.merge(fte_df, usd_cost_df, on="Month", how="outer").fillna(0)
    monthly = monthly.sort_values("Month")
    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")

    return monthly


def create_simulation_version_from_baseline(
    baseline_monthly: pd.DataFrame,
) -> pd.DataFrame:
    sim = baseline_monthly.copy()
    sim["SimulationFTE"] = sim["BaselineFTE"]
    sim["SimulationUSD"] = sim["BaselineUSD"]
    return sim


def apply_uploaded_drivers_to_baseline_methodology(
    monthly_df: pd.DataFrame,
    drivers_df: pd.DataFrame,
    rules_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    Placeholder methodology layer.

    Baseline is never changed.
    The uploaded Drivers and Calculation_Method files are retained and available.
    Simulation starts from already loaded Baseline values.
    """
    result = monthly_df.copy()
    result["MethodologySource"] = "Calculation_Method + Drivers"
    return result


# ============================================================
# OpenAI simulation parser
# ============================================================

def get_openai_api_key() -> Optional[str]:
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass

    return os.getenv("OPENAI_API_KEY")


def call_openai_api(user_instruction: str) -> Dict[str, Any]:
    api_key = get_openai_api_key()

    if not api_key:
        return fallback_parse_instruction(user_instruction)

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)

        system_prompt = """
You are a workforce planning simulation engine.

Return valid JSON only.

Important rules:
1. Never change Baseline.
2. Use Baseline as starting point for Simulation.
3. Baseline methodology is defined by Calculation_Method rules per GL Account.
4. Baseline driver values come from uploaded Drivers file.
5. Create changes only for the Simulation version.

JSON format:
{
  "summary": "short summary",
  "actions": [
    {
      "action_type": "hire | layoff | salary_change | cost_change | driver_change",
      "fte_delta": number,
      "cost_pct_delta": number,
      "cost_abs_delta": number,
      "driver": string or null,
      "effective_month": "YYYYMM"
    }
  ]
}

Use YYYYMM month format.
Use 202601 if no month is provided.
Percentages must be decimals, e.g. 5% = 0.05.
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_instruction},
            ],
        )

        return json.loads(response.choices[0].message.content.strip())

    except Exception:
        return fallback_parse_instruction(user_instruction)


def fallback_parse_instruction(text: str) -> Dict[str, Any]:
    original = text
    text = text.lower()

    number = extract_first_number(text)
    month = extract_month_yyyymm(text)

    action = {
        "action_type": "cost_change",
        "fte_delta": 0,
        "cost_pct_delta": 0,
        "cost_abs_delta": 0,
        "driver": None,
        "effective_month": month,
    }

    if any(w in text for w in ["hire", "add fte", "increase fte", "recruit"]):
        action["action_type"] = "hire"
        action["fte_delta"] = abs(number)

    elif any(w in text for w in ["layoff", "remove fte", "reduce fte", "cut fte"]):
        action["action_type"] = "layoff"
        action["fte_delta"] = -abs(number)

    elif any(w in text for w in ["salary", "cost", "labor cost", "usd", "wage"]):
        action["action_type"] = "cost_change"
        action["cost_pct_delta"] = abs(number) / 100

        if any(w in text for w in ["reduce", "decrease", "cut", "lower"]):
            action["cost_pct_delta"] = -abs(action["cost_pct_delta"])

    elif "driver" in text:
        action["action_type"] = "driver_change"
        action["driver"] = "Manual Driver"
        action["cost_pct_delta"] = number / 100

    return {"summary": original[:100], "actions": [action]}


def extract_first_number(text: str) -> float:
    match = re.search(r"[-+]?\d*\.?\d+", text)
    return float(match.group()) if match else 0.0


def extract_month_yyyymm(text: str) -> str:
    direct_match = re.search(r"\b(20\d{2})(0[1-9]|1[0-2])\b", text)
    if direct_match:
        return direct_match.group()

    month_map = {
        "jan": "202601", "january": "202601",
        "feb": "202602", "february": "202602",
        "mar": "202603", "march": "202603",
        "apr": "202604", "april": "202604",
        "may": "202605",
        "jun": "202606", "june": "202606",
        "jul": "202607", "july": "202607",
        "aug": "202608", "august": "202608",
        "sep": "202609", "september": "202609",
        "oct": "202610", "october": "202610",
        "nov": "202611", "november": "202611",
        "dec": "202612", "december": "202612",
    }

    for key, value in month_map.items():
        if key in text:
            return value

    return "202601"


# ============================================================
# Simulation logic
# ============================================================

def apply_simulation_logic(
    baseline_monthly: pd.DataFrame,
    events: List[Dict[str, Any]],
) -> pd.DataFrame:
    df = create_simulation_version_from_baseline(baseline_monthly)

    avg_usd_per_fte = (
        df["BaselineUSD"].sum() / df["BaselineFTE"].sum()
        if df["BaselineFTE"].sum() != 0
        else 0
    )

    for event in events:
        for action in event.get("actions", []):
            effective_month = pd.to_datetime(
                str(action.get("effective_month", "202601")),
                format="%Y%m",
                errors="coerce",
            )

            if pd.isna(effective_month):
                effective_month = df["Month"].min()

            mask = df["Month"] >= effective_month

            action_type = action.get("action_type")
            fte_delta = float(action.get("fte_delta", 0) or 0)
            cost_pct_delta = float(action.get("cost_pct_delta", 0) or 0)
            cost_abs_delta = float(action.get("cost_abs_delta", 0) or 0)

            if action_type in ["hire", "layoff"]:
                df.loc[mask, "SimulationFTE"] += fte_delta
                df.loc[mask, "SimulationUSD"] += fte_delta * avg_usd_per_fte

            elif action_type in ["salary_change", "cost_change", "driver_change"]:
                df.loc[mask, "SimulationUSD"] *= 1 + cost_pct_delta
                df.loc[mask, "SimulationUSD"] += cost_abs_delta

    df["SimulationFTE"] = df["SimulationFTE"].clip(lower=0)
    df["SimulationUSD"] = df["SimulationUSD"].clip(lower=0)

    return df


# ============================================================
# Charts and summary
# ============================================================

def render_chart(
    title: str,
    df: pd.DataFrame,
    baseline_metric: str,
    simulation_metric: str,
    y_title: str,
    height: int,
) -> go.Figure:
    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=df["MonthLabel"],
            y=df[baseline_metric],
            name="Baseline",
            marker_color="#2563EB",
        )
    )

    fig.add_trace(
        go.Bar(
            x=df["MonthLabel"],
            y=df[simulation_metric],
            name="Simulation",
            marker_color="#F97316",
        )
    )

    fig.update_layout(
        title=title,
        barmode="group",
        height=height,
        plot_bgcolor="#111827",
        paper_bgcolor="#111827",
        font=dict(color="#F9FAFB"),
        margin=dict(l=30, r=30, t=62, b=90),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1.0,
        ),
        yaxis_title=y_title,
    )

    fig.update_yaxes(gridcolor="#374151")
    fig.update_xaxes(
        showgrid=False,
        tickangle=-45,
        tickmode="array",
        tickvals=df["MonthLabel"].tolist(),
        ticktext=df["MonthLabel"].tolist(),
    )

    return fig


def build_summary(df: pd.DataFrame, events: List[Dict[str, Any]]) -> List[str]:
    fte_impact = df["SimulationFTE"].sum() - df["BaselineFTE"].sum()
    cost_impact = df["SimulationUSD"].sum() - df["BaselineUSD"].sum()

    summary = [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total USD labor cost impact: ${cost_impact:+,.0f}",
        f"Baseline preserved: Yes",
        f"Simulation changes applied: {len(events)}",
    ]

    if events:
        summary.append(f"Latest change: {events[-1].get('summary', '')}")

    return [item[:100] for item in summary[:6]]


# ============================================================
# UI
# ============================================================

def apply_css():
    st.markdown(
        """
        <style>
        .stApp {
            background-color: #0B0F17;
            color: #F9FAFB;
        }

        .block-container {
            padding-top: 1rem;
            padding-left: 2rem;
            padding-right: 2rem;
            max-width: 100%;
        }

        .main-header {
            background: linear-gradient(90deg, #0F172A, #1E293B);
            color: white;
            padding: 16px 24px;
            border-radius: 14px;
            margin-bottom: 20px;
            display: grid;
            grid-template-columns: 1fr 2fr 1fr;
            align-items: center;
        }

        .logo {
            width: 58px;
            height: 38px;
            border-radius: 8px;
            background: white;
            color: #0F766E;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 900;
        }

        .title {
            text-align: center;
            font-size: 26px;
            font-weight: 800;
        }

        .user {
            display: flex;
            justify-content: flex-end;
            gap: 12px;
            align-items: center;
            font-weight: 700;
        }

        .avatar {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            border: 2px dashed #94A3B8;
        }

        div[data-testid="stMetric"] {
            background: #111827;
            padding: 18px 16px;
            border-radius: 10px;
            border: 1px solid #1F2937;
            min-height: 126px;
            height: 126px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            align-items: flex-start;
        }

        div[data-testid="stMetricLabel"] {
            color: #F9FAFB;
            height: 24px;
            line-height: 24px;
            margin-top: 0px;
        }

        div[data-testid="stMetricValue"] {
            color: #F9FAFB;
            font-size: 28px;
            line-height: 34px;
            margin-top: 12px;
        }

        div[data-testid="stMetricDelta"] {
            margin-top: 8px;
        }

        .summary {
            background: #111827;
            border: 1px solid #334155;
            border-left: 6px solid #F97316;
            padding: 18px 24px;
            border-radius: 14px;
            margin-top: 18px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header():
    st.markdown(
        f"""
        <div class="main-header">
            <div class="logo">TEVA</div>
            <div class="title">{REPORT_TITLE}</div>
            <div class="user">
                <div>TEST USER</div>
                <div class="avatar"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_excel_help():
    with st.expander("Expected Excel structures"):
        st.markdown(
            """
            **1. Excel file: `data_sample_anonymized`**

            Required sheet: **FTE & Costs**

            Required row columns:
            - Company Code
            - Cost Center
            - Profit Center
            - Business Area
            - Segment
            - Employee
            - Account

            Supported extra row column:
            - Version

            Monthly values must be in columns using **YYYYMM** format:
            - 202601
            - 202602
            - 202603

            Baseline rows must have:
            - Version = Baseline

            FTE chart uses:
            - Account containing `FTE`
            - Version = Baseline

            USD labor cost chart uses:
            - Account containing `USD`
            - Version = Baseline

            **2. Excel file: `Drivers`**

            Required sheet: **SAC Driver**

            Required columns:
            - Driver
            - Monthly date columns in **YYYYMM** format

            **3. Excel file: `Calculation_method`**

            Required sheet: **NDC Per country Rules**

            Required columns:
            - Description
            - Global Account
            - SAC Driver
            - NDC Comment
            - Israel
            """
        )


def initialize_state():
    if "data_sample_df" not in st.session_state:
        st.session_state.data_sample_df = generate_mock_source_data()

    if "drivers_df" not in st.session_state:
        st.session_state.drivers_df = generate_mock_drivers()

    if "rules_df" not in st.session_state:
        st.session_state.rules_df = generate_mock_rules()

    if "simulation_events" not in st.session_state:
        st.session_state.simulation_events = []

    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [
            {
                "role": "assistant",
                "content": (
                    "Enter scenario, e.g. hire 10 FTE from 202604 "
                    "or reduce USD labor cost by 5%."
                ),
            }
        ]

    if "data_sample_source" not in st.session_state:
        st.session_state.data_sample_source = "Mock data"

    if "drivers_source" not in st.session_state:
        st.session_state.drivers_source = "Mock drivers"

    if "rules_source" not in st.session_state:
        st.session_state.rules_source = "Mock rules"


def main():
    initialize_state()
    apply_css()
    render_header()

    baseline_monthly = source_to_monthly_baseline(st.session_state.data_sample_df)
    baseline_monthly = apply_uploaded_drivers_to_baseline_methodology(
        baseline_monthly,
        st.session_state.drivers_df,
        st.session_state.rules_df,
    )

    simulation_df = apply_simulation_logic(
        baseline_monthly,
        st.session_state.simulation_events,
    )

    baseline_total_fte = simulation_df["BaselineFTE"].sum()
    simulation_total_fte = simulation_df["SimulationFTE"].sum()
    baseline_total_usd = simulation_df["BaselineUSD"].sum()
    simulation_total_usd = simulation_df["SimulationUSD"].sum()

    left, right = st.columns([1, 4], gap="large")

    with left:
        st.subheader("GPT Simulation")

        data_file = st.file_uploader(
            "1. Upload data_sample_anonymized",
            type=["xlsx"],
            key="data_sample_uploader",
        )

        if data_file:
            try:
                st.session_state.data_sample_df = load_data_sample(data_file)
                st.session_state.data_sample_source = data_file.name
                st.session_state.simulation_events = []
                st.success("Data sample loaded.")
                st.rerun()
            except Exception as exc:
                st.warning(str(exc))

        drivers_file = st.file_uploader(
            "2. Upload Drivers",
            type=["xlsx"],
            key="drivers_uploader",
        )

        if drivers_file:
            try:
                st.session_state.drivers_df = load_drivers(drivers_file)
                st.session_state.drivers_source = drivers_file.name
                st.success("Drivers loaded.")
                st.rerun()
            except Exception as exc:
                st.warning(str(exc))

        calc_file = st.file_uploader(
            "3. Upload Calculation_method",
            type=["xlsx"],
            key="calculation_method_uploader",
        )

        if calc_file:
            try:
                st.session_state.rules_df = load_calculation_rules(calc_file)
                st.session_state.rules_source = calc_file.name
                st.success("Calculation method loaded.")
                st.rerun()
            except Exception as exc:
                st.warning(str(exc))

        render_excel_help()

        st.caption(f"Data: {st.session_state.data_sample_source}")
        st.caption(f"Drivers: {st.session_state.drivers_source}")
        st.caption(f"Rules: {st.session_state.rules_source}")

        st.divider()

        for msg in st.session_state.chat_messages:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        prompt = st.chat_input("Enter simulation instruction...")

        if prompt:
            st.session_state.chat_messages.append(
                {"role": "user", "content": prompt}
            )

            parsed = call_openai_api(prompt)
            st.session_state.simulation_events.append(parsed)

            st.session_state.chat_messages.append(
                {
                    "role": "assistant",
                    "content": (
                        f"Applied to Simulation only: "
                        f"{parsed.get('summary', 'Simulation change')}"
                    ),
                }
            )

            st.rerun()

        if st.button("Reset simulation", use_container_width=True):
            st.session_state.simulation_events = []
            st.session_state.chat_messages = [
                {
                    "role": "assistant",
                    "content": "Simulation reset. Baseline remains unchanged.",
                }
            ]
            st.rerun()

    with right:
        c1, c2, c3, c4 = st.columns(4)

        c1.metric("Baseline FTE", f"{baseline_total_fte:,.1f}")
        c2.metric(
            "Simulation FTE",
            f"{simulation_total_fte:,.1f}",
            delta=f"{simulation_total_fte - baseline_total_fte:+,.1f}",
        )
        c3.metric("Baseline USD Labor Cost", f"${baseline_total_usd:,.0f}")
        c4.metric(
            "Simulation USD Labor Cost",
            f"${simulation_total_usd:,.0f}",
            delta=f"${simulation_total_usd - baseline_total_usd:+,.0f}",
        )

        st.plotly_chart(
            render_chart(
                "FTE over Time",
                simulation_df,
                "BaselineFTE",
                "SimulationFTE",
                "FTE",
                height=240,
            ),
            use_container_width=True,
        )

        st.plotly_chart(
            render_chart(
                "Total Cost of Labor over Time",
                simulation_df,
                "BaselineUSD",
                "SimulationUSD",
                "USD Labor Cost",
                height=480,
            ),
            use_container_width=True,
        )

        st.markdown('<div class="summary">', unsafe_allow_html=True)
        st.subheader("Simulation Impact Summary")

        for item in build_summary(
            simulation_df,
            st.session_state.simulation_events,
        ):
            st.markdown(f"- {item}")

        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("Preview loaded data"):
            tab1, tab2, tab3, tab4 = st.tabs(
                ["FTE & Costs", "SAC Driver", "NDC Rules", "Monthly Output"]
            )

            with tab1:
                st.dataframe(st.session_state.data_sample_df.head(50))

            with tab2:
                st.dataframe(st.session_state.drivers_df.head(50))

            with tab3:
                st.dataframe(st.session_state.rules_df.head(50))

            with tab4:
                st.dataframe(simulation_df)


if __name__ == "__main__":
    main()
