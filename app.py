# app.py
import os
import json
import re
from datetime import datetime
from typing import Dict, List, Any, Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


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
    accounts = ["FTE", "Salary", "Bonus", "Benefits", "Payroll Tax"]
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
                elif account == "Salary":
                    row[col] = 5200 + i * 50
                elif account == "Bonus":
                    row[col] = 400
                elif account == "Benefits":
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
            "Bonus",
            "Benefits",
            "Payroll Tax",
            "Hiring Plan",
            "Layoff Plan",
        ]
    }

    for month in months:
        data[month.strftime("%Y%m")] = [0.03, 0.10, 0.02, 0.12, 0, 0]

    return pd.DataFrame(data)


def generate_mock_rules() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Description": [
                "FTE baseline account",
                "Salary cost calculated from salary driver",
                "Bonus cost calculated from bonus driver",
                "Benefits cost calculated from benefits driver",
                "Tax cost calculated from payroll tax driver",
            ],
            "Global Account": ["FTE", "Salary", "Bonus", "Benefits", "Payroll Tax"],
            "SAC Driver": [
                "Hiring Plan",
                "Salary Increase",
                "Bonus",
                "Benefits",
                "Payroll Tax",
            ],
            "NDC Comment": [
                "Baseline FTE volume",
                "Labor salary cost",
                "Labor bonus cost",
                "Labor benefits cost",
                "Labor tax cost",
            ],
            "Israel": ["Yes", "Yes", "Yes", "Yes", "Yes"],
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


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={col: clean_column_name(col) for col in df.columns})


def is_yyyymm_column(col: Any) -> bool:
    return bool(re.fullmatch(r"\d{6}", clean_column_name(col)))


def validate_required_columns(
    df: pd.DataFrame,
    required: List[str],
    sheet_name: str,
) -> None:
    existing = {str(col).strip().lower(): col for col in df.columns}
    missing = [col for col in required if col.strip().lower() not in existing]

    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}"
        )


def get_actual_column(df: pd.DataFrame, expected_name: str) -> Optional[str]:
    for col in df.columns:
        if str(col).strip().lower() == expected_name.strip().lower():
            return col
    return None


def get_month_columns(df: pd.DataFrame, excluded: List[str]) -> List[str]:
    excluded_lower = [x.lower() for x in excluded]

    month_cols = [
        col for col in df.columns
        if str(col).strip().lower() not in excluded_lower and is_yyyymm_column(col)
    ]

    return sorted(month_cols, key=lambda x: pd.to_datetime(str(x), format="%Y%m"))


def validate_month_columns(
    df: pd.DataFrame,
    excluded: List[str],
    sheet_name: str,
) -> List[str]:
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
    df = normalize_columns(df)

    validate_required_columns(df, DATA_SAMPLE_ROW_COLUMNS, DATA_SAMPLE_SHEET)

    if get_actual_column(df, "Version") is None:
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
    df = normalize_columns(df)

    validate_required_columns(df, ["Driver"], DRIVERS_SHEET)

    month_cols = validate_month_columns(df, ["Driver"], DRIVERS_SHEET)

    for col in month_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_calculation_rules(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, CALC_RULES_SHEET)
    df = normalize_columns(df)

    validate_required_columns(df, CALC_RULES_REQUIRED_COLUMNS, CALC_RULES_SHEET)

    return df


# ============================================================
# Baseline replacement
# ============================================================

def replace_baseline_with_uploaded_data(uploaded_file) -> None:
    new_df = load_data_sample(uploaded_file)

    st.session_state.data_sample_df = new_df
    st.session_state.data_sample_source = uploaded_file.name
    st.session_state.simulation_events = []
    st.session_state.audit_log = []
    st.session_state.simulation_requests = []


def reset_to_mock_baseline() -> None:
    st.session_state.data_sample_df = generate_mock_source_data()
    st.session_state.data_sample_source = "Mock data"
    st.session_state.simulation_events = []
    st.session_state.audit_log = []
    st.session_state.simulation_requests = []
    st.session_state.last_data_sample_upload_id = None


# ============================================================
# Baseline logic
# ============================================================

def filter_version(df: pd.DataFrame, version: str) -> pd.DataFrame:
    version_col = get_actual_column(df, "Version")

    if version_col is None:
        return df.copy()

    filtered = df[
        df[version_col].astype(str).str.strip().str.lower()
        == version.strip().lower()
    ].copy()

    if filtered.empty:
        st.warning(
            f"No rows found for Version = '{version}'. "
            "Using all rows from uploaded file instead."
        )
        return df.copy()

    return filtered


def identify_fte_rows(df: pd.DataFrame) -> pd.Series:
    account_col = get_actual_column(df, "Account")

    if account_col is None:
        return pd.Series(False, index=df.index)

    account_text = df[account_col].astype(str).str.lower().str.strip()

    return (
        account_text.str.contains("fte", na=False)
        | account_text.str.contains("headcount", na=False)
        | account_text.str.fullmatch("hc", na=False)
    )


def identify_cost_rows(df: pd.DataFrame) -> pd.Series:
    return ~identify_fte_rows(df)


def source_to_monthly_baseline(source_df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(source_df)
    baseline_df = filter_version(df, "Baseline")

    version_col = get_actual_column(baseline_df, "Version") or "Version"

    id_columns = DATA_SAMPLE_ROW_COLUMNS.copy()
    if version_col in baseline_df.columns:
        id_columns.append(version_col)

    month_cols = get_month_columns(
        baseline_df,
        DATA_SAMPLE_ROW_COLUMNS + ["Version"],
    )

    if not month_cols:
        raise ValueError("No YYYYMM month columns found in uploaded data.")

    long_df = baseline_df.melt(
        id_vars=id_columns,
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

    cost_df = (
        long_df[identify_cost_rows(long_df)]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "BaselineUSD"})
    )

    monthly = pd.merge(fte_df, cost_df, on="Month", how="outer").fillna(0)
    monthly = monthly.sort_values("Month")
    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")

    return monthly


def create_simulation_from_baseline(baseline_monthly: pd.DataFrame) -> pd.DataFrame:
    df = baseline_monthly.copy()
    df["SimulationFTE"] = df["BaselineFTE"]
    df["SimulationUSD"] = df["BaselineUSD"]
    return df


# ============================================================
# GPT parser
# ============================================================

def build_agent_rules_prompt() -> str:
    return """
You are a Workforce Planning Simulation Agent.

Rules:
1. Baseline is read-only.
2. Simulation starts as a copy of Baseline.
3. All changes apply only to Simulation.
4. Simulation changes are cumulative until reset.
5. Effective date is mandatory.
6. If effective date is missing, return clarification_required.
7. Changes affect all months from effective date forward unless end date is provided.
8. Hire, recruit, staff increase, add employees, open cost center, open site increase FTE and cost.
9. Layoff, workforce reduction, staff reduction, close cost center, close site decrease FTE and cost.
10. FTE and cost cannot become negative.
11. Merit, salary, bonus, payroll tax, benefit changes affect cost only unless FTE is explicitly mentioned.
12. Driver changes must recalculate linked GL accounts.
13. Driver-to-account mapping comes from Calculation Method.
14. Never guess site, country, cost center, or employee population.

Return valid JSON only.

Schema:
{
  "status": "ready_to_apply" | "clarification_required",
  "clarification_question": "string or null",
  "user_request": "string",
  "assumptions": ["string"],
  "effective_date": "YYYYMM or null",
  "end_date": "YYYYMM or null",
  "changes_applied": ["string"],
  "drivers_impacted": ["string"],
  "gl_accounts_impacted": ["string"],
  "fte_impact_description": "string",
  "labor_cost_impact_description": "string",
  "monthly_simulation_results_description": "string",
  "executive_summary": "string",
  "actions": [
    {
      "action_type": "hire | layoff | salary_change | cost_change | driver_change | organization_change",
      "fte_delta": number,
      "cost_pct_delta": number,
      "cost_abs_delta": number,
      "driver": "string or null",
      "gl_account": "string or null",
      "scope": "string or null",
      "effective_month": "YYYYMM",
      "end_month": "YYYYMM or null"
    }
  ]
}
"""


def summarize_dataframe_for_prompt(df: pd.DataFrame, max_rows: int = 20) -> str:
    if df is None or df.empty:
        return "No data loaded."

    preview = df.head(max_rows).to_dict(orient="records")
    return json.dumps(preview, default=str)[:8000]


def get_openai_api_key() -> Optional[str]:
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass

    return os.getenv("OPENAI_API_KEY")


def call_openai_api(
    user_instruction: str,
    baseline_monthly: pd.DataFrame,
    drivers_df: pd.DataFrame,
    rules_df: pd.DataFrame,
) -> Dict[str, Any]:
    api_key = get_openai_api_key()

    if not api_key:
        return fallback_parse_instruction(user_instruction)

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)

        context_prompt = f"""
Baseline monthly summary:
{summarize_dataframe_for_prompt(baseline_monthly)}

Drivers preview:
{summarize_dataframe_for_prompt(drivers_df)}

Calculation Method preview:
{summarize_dataframe_for_prompt(rules_df)}

User simulation instruction:
{user_instruction}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system", "content": build_agent_rules_prompt()},
                {"role": "user", "content": context_prompt},
            ],
        )

        return json.loads(response.choices[0].message.content.strip())

    except Exception:
        return fallback_parse_instruction(user_instruction)


def fallback_parse_instruction(text: str) -> Dict[str, Any]:
    original = text
    lower = text.lower()

    has_effective_date = bool(re.search(r"\b(20\d{2})(0[1-9]|1[0-2])\b", lower)) or any(
        month in lower
        for month in [
            "jan", "feb", "mar", "apr", "may", "jun",
            "jul", "aug", "sep", "oct", "nov", "dec",
        ]
    )

    if not has_effective_date:
        return {
            "status": "clarification_required",
            "clarification_question": "What is the effective date? Please use YYYYMM format.",
            "user_request": original,
            "assumptions": [],
            "effective_date": None,
            "end_date": None,
            "changes_applied": [],
            "drivers_impacted": [],
            "gl_accounts_impacted": [],
            "fte_impact_description": "Not calculated.",
            "labor_cost_impact_description": "Not calculated.",
            "monthly_simulation_results_description": "Not calculated.",
            "executive_summary": "Clarification required before simulation can be calculated.",
            "actions": [],
        }

    number = extract_first_number(lower)
    month = extract_month_yyyymm(lower)

    action = {
        "action_type": "cost_change",
        "fte_delta": 0,
        "cost_pct_delta": 0,
        "cost_abs_delta": 0,
        "driver": None,
        "gl_account": None,
        "scope": None,
        "effective_month": month,
        "end_month": None,
    }

    changes = []

    if any(w in lower for w in ["hire", "add fte", "increase fte", "recruit", "staff increase"]):
        action["action_type"] = "hire"
        action["fte_delta"] = abs(number)
        changes.append(f"Increase Simulation FTE by {abs(number):,.2f} from {month}.")

    elif any(w in lower for w in ["layoff", "remove fte", "reduce fte", "cut fte", "workforce reduction"]):
        action["action_type"] = "layoff"
        action["fte_delta"] = -abs(number)
        changes.append(f"Decrease Simulation FTE by {abs(number):,.2f} from {month}.")

    elif any(w in lower for w in ["salary", "cost", "labor cost", "usd", "wage", "merit", "bonus", "benefit", "payroll tax"]):
        action["action_type"] = "cost_change"
        action["cost_pct_delta"] = abs(number) / 100

        if any(w in lower for w in ["reduce", "decrease", "cut", "lower", "reduction"]):
            action["cost_pct_delta"] = -abs(action["cost_pct_delta"])

        changes.append(f"Change Simulation cost by {action['cost_pct_delta']:+.2%} from {month}.")

    return {
        "status": "ready_to_apply",
        "clarification_question": None,
        "user_request": original,
        "assumptions": [
            "Baseline remains unchanged.",
            "Simulation starts as a copy of Baseline.",
            "Change applies from effective month forward.",
        ],
        "effective_date": month,
        "end_date": None,
        "changes_applied": changes,
        "drivers_impacted": [],
        "gl_accounts_impacted": [],
        "fte_impact_description": "Calculated after applying simulation action.",
        "labor_cost_impact_description": "Calculated after applying simulation action.",
        "monthly_simulation_results_description": "Updated monthly Simulation values are shown in charts.",
        "executive_summary": original[:100],
        "actions": [action],
    }


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
    df = create_simulation_from_baseline(baseline_monthly)

    avg_usd_per_fte = (
        df["BaselineUSD"].sum() / df["BaselineFTE"].sum()
        if df["BaselineFTE"].sum() != 0
        else 0
    )

    for event in events:
        if event.get("status") != "ready_to_apply":
            continue

        for action in event.get("actions", []):
            effective_month = pd.to_datetime(
                str(action.get("effective_month", "202601")),
                format="%Y%m",
                errors="coerce",
            )

            end_month = pd.to_datetime(
                str(action.get("end_month")),
                format="%Y%m",
                errors="coerce",
            )

            if pd.isna(effective_month):
                continue

            mask = df["Month"] >= effective_month

            if pd.notna(end_month):
                mask &= df["Month"] <= end_month

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


def create_audit_entry(
    event: Dict[str, Any],
    before_df: pd.DataFrame,
    after_df: pd.DataFrame,
) -> Dict[str, Any]:
    fte_impact = after_df["SimulationFTE"].sum() - before_df["SimulationFTE"].sum()
    cost_impact = after_df["SimulationUSD"].sum() - before_df["SimulationUSD"].sum()

    return {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "User Request": event.get("user_request", ""),
        "Effective Date": event.get("effective_date", ""),
        "Assumptions": "; ".join(event.get("assumptions", [])),
        "FTE Impact": fte_impact,
        "Cost Impact": cost_impact,
        "Executive Summary": event.get("executive_summary", ""),
    }


# ============================================================
# Charts, metrics, summary
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


def render_custom_metric(
    title: str,
    main_value: str,
    sub_label: str = "",
    sub_value: str = "",
    delta: str = "",
    delta_positive: bool = True,
) -> None:
    delta_class = "positive" if delta_positive else "negative"

    delta_html = (
        f"<div class='metric-delta {delta_class}'>{delta}</div>"
        if delta
        else ""
    )

    sub_html = (
        f"<div class='metric-sub'>{sub_label}: <b>{sub_value}</b></div>"
        if sub_label
        else ""
    )

    st.markdown(
        f"""
        <div class="custom-metric">
            <div class="metric-title">{title}</div>
            <div class="metric-main-row">
                <div class="metric-main">{main_value}</div>
                {delta_html}
            </div>
            {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_summary(df: pd.DataFrame, events: List[Dict[str, Any]]) -> List[str]:
    fte_impact = df["SimulationFTE"].sum() - df["BaselineFTE"].sum()
    cost_impact = df["SimulationUSD"].sum() - df["BaselineUSD"].sum()

    summary = [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total labor cost impact: ${cost_impact:+,.0f}",
        f"Baseline preserved: Yes",
        f"Simulation changes applied: {len([e for e in events if e.get('status') == 'ready_to_apply'])}",
    ]

    if events:
        summary.append(f"Latest: {events[-1].get('executive_summary', events[-1].get('user_request', ''))}")

    return [item[:100] for item in summary[:6]]


def format_agent_response(event: Dict[str, Any]) -> str:
    if event.get("status") == "clarification_required":
        return event.get("clarification_question", "Clarification required.")

    return (
        f"Executive Summary: {event.get('executive_summary', '')}<br>"
        f"Effective Date: {event.get('effective_date', '')}<br>"
        f"Changes Applied: {', '.join(event.get('changes_applied', [])) or 'None'}"
    )


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

        .metric-row-spacer {
            height: 5px;
        }

        .custom-metric {
            background: #111827;
            padding: 18px 16px;
            border-radius: 10px;
            border: 1px solid #1F2937;
            height: 126px;
            min-height: 126px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            align-items: stretch;
            box-sizing: border-box;
        }

        .metric-title {
            color: #F9FAFB;
            font-size: 14px;
            line-height: 24px;
            height: 24px;
            font-weight: 600;
        }

        .metric-main-row {
            display: flex;
            align-items: baseline;
            justify-content: space-between;
            gap: 8px;
            width: 100%;
            margin-top: 12px;
        }

        .metric-main {
            color: #F9FAFB;
            font-size: 28px;
            line-height: 34px;
            font-weight: 700;
            white-space: nowrap;
        }

        .metric-sub {
            color: #CBD5E1;
            font-size: 10px;
            line-height: 12px;
            margin-top: 8px;
        }

        .metric-delta {
            font-size: 13px;
            line-height: 16px;
            font-weight: 800;
            text-align: right;
            white-space: nowrap;
            margin-left: auto;
        }

        .metric-delta.positive {
            color: #22C55E;
        }

        .metric-delta.negative {
            color: #EF4444;
        }

        .simulation-input-box {
            background: #111827;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 14px;
            margin-bottom: 14px;
        }

        .simulation-history-box {
            background: #111827;
            border: 1px solid #334155;
            border-radius: 12px;
            padding: 12px;
            margin-bottom: 10px;
        }

        .simulation-history-meta {
            color: #94A3B8;
            font-size: 11px;
            margin-bottom: 6px;
        }

        .simulation-history-text {
            color: #F9FAFB;
            font-size: 13px;
            line-height: 1.35;
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
            **Data Sample file**

            Required sheet: **FTE & Costs**

            Required row columns:
            - Company Code
            - Cost Center
            - Profit Center
            - Business Area
            - Segment
            - Employee
            - Account

            Supported extra column:
            - Version

            Monthly values must be in columns using **YYYYMM** format.
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

    if "simulation_requests" not in st.session_state:
        st.session_state.simulation_requests = []

    if "audit_log" not in st.session_state:
        st.session_state.audit_log = []

    if "data_sample_source" not in st.session_state:
        st.session_state.data_sample_source = "Mock data"

    if "drivers_source" not in st.session_state:
        st.session_state.drivers_source = "Mock drivers"

    if "rules_source" not in st.session_state:
        st.session_state.rules_source = "Mock rules"

    if "last_data_sample_upload_id" not in st.session_state:
        st.session_state.last_data_sample_upload_id = None

    if "last_drivers_upload_id" not in st.session_state:
        st.session_state.last_drivers_upload_id = None

    if "last_rules_upload_id" not in st.session_state:
        st.session_state.last_rules_upload_id = None

    if "simulation_input_counter" not in st.session_state:
        st.session_state.simulation_input_counter = 0


# ============================================================
# Main app
# ============================================================

def main():
    initialize_state()
    apply_css()
    render_header()

    left, right = st.columns([1, 4], gap="large")

    with left:
        st.subheader("GPT Simulation")

        data_file = st.file_uploader(
            "Upload / Reupload Data Sample",
            type=["xlsx"],
            key="data_sample_reupload",
            help="Reuploading replaces the current baseline and clears simulation.",
        )

        if data_file is not None:
            upload_id = f"{data_file.name}_{data_file.size}"

            if upload_id != st.session_state.last_data_sample_upload_id:
                try:
                    replace_baseline_with_uploaded_data(data_file)
                    st.session_state.last_data_sample_upload_id = upload_id
                    st.success("Baseline replaced with uploaded Data Sample.")
                    st.rerun()
                except Exception as exc:
                    st.warning(str(exc))

        if st.button("Reset baseline to mock data", use_container_width=True):
            reset_to_mock_baseline()
            st.rerun()

        drivers_file = st.file_uploader(
            "Upload Drivers",
            type=["xlsx"],
            key="drivers_uploader",
        )

        if drivers_file is not None:
            upload_id = f"{drivers_file.name}_{drivers_file.size}"

            if upload_id != st.session_state.last_drivers_upload_id:
                try:
                    st.session_state.drivers_df = load_drivers(drivers_file)
                    st.session_state.drivers_source = drivers_file.name
                    st.session_state.last_drivers_upload_id = upload_id
                    st.success("Drivers loaded.")
                    st.rerun()
                except Exception as exc:
                    st.warning(str(exc))

        calc_file = st.file_uploader(
            "Upload Calculation_method",
            type=["xlsx"],
            key="calculation_method_uploader",
        )

        if calc_file is not None:
            upload_id = f"{calc_file.name}_{calc_file.size}"

            if upload_id != st.session_state.last_rules_upload_id:
                try:
                    st.session_state.rules_df = load_calculation_rules(calc_file)
                    st.session_state.rules_source = calc_file.name
                    st.session_state.last_rules_upload_id = upload_id
                    st.success("Calculation method loaded.")
                    st.rerun()
                except Exception as exc:
                    st.warning(str(exc))

        render_excel_help()

        st.caption(f"Data Sample: {st.session_state.data_sample_source}")
        st.caption(f"Drivers: {st.session_state.drivers_source}")
        st.caption(f"Rules: {st.session_state.rules_source}")

    try:
        baseline_monthly = source_to_monthly_baseline(st.session_state.data_sample_df)
        current_simulation_df = apply_simulation_logic(
            baseline_monthly,
            st.session_state.simulation_events,
        )
    except Exception as exc:
        with right:
            st.error(f"Data preparation error: {exc}")
        return

    with left:
        st.divider()

        st.markdown('<div class="simulation-input-box">', unsafe_allow_html=True)

        input_key = f"simulation_input_text_{st.session_state.simulation_input_counter}"

        simulation_prompt = st.text_area(
            "Simulation input",
            placeholder="Example: Hire 10 FTE effective 202604",
            height=110,
            key=input_key,
        )

        run_simulation = st.button("Run simulation", use_container_width=True)

        st.markdown("</div>", unsafe_allow_html=True)

        if run_simulation and simulation_prompt.strip():
            prompt = simulation_prompt.strip()
            before_df = current_simulation_df.copy()

            parsed = call_openai_api(
                prompt,
                baseline_monthly,
                st.session_state.drivers_df,
                st.session_state.rules_df,
            )

            if parsed.get("status") == "ready_to_apply":
                st.session_state.simulation_events.append(parsed)

                after_df = apply_simulation_logic(
                    baseline_monthly,
                    st.session_state.simulation_events,
                )

                st.session_state.audit_log.append(
                    create_audit_entry(parsed, before_df, after_df)
                )

            st.session_state.simulation_requests.insert(
                0,
                {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "request": prompt,
                    "response": format_agent_response(parsed),
                    "status": parsed.get("status", ""),
                },
            )

            st.session_state.simulation_input_counter += 1
            st.rerun()

        if st.button("Reset simulation only", use_container_width=True):
            st.session_state.simulation_events = []
            st.session_state.audit_log = []
            st.session_state.simulation_requests = []
            st.session_state.simulation_input_counter += 1
            st.rerun()

        st.subheader("Executed simulations")

        if st.session_state.simulation_requests:
            for item in st.session_state.simulation_requests:
                st.markdown(
                    f"""
                    <div class="simulation-history-box">
                        <div class="simulation-history-meta">
                            {item["timestamp"]} | {item["status"]}
                        </div>
                        <div class="simulation-history-text">
                            <b>Request:</b> {item["request"]}<br><br>
                            <b>Result:</b> {item["response"]}
                        </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
        else:
            st.caption("No simulations executed yet.")

    simulation_df = apply_simulation_logic(
        baseline_monthly,
        st.session_state.simulation_events,
    )

    baseline_avg_fte = simulation_df["BaselineFTE"].mean()
    simulation_avg_fte = simulation_df["SimulationFTE"].mean()

    baseline_end_fte = simulation_df["BaselineFTE"].iloc[-1]
    simulation_end_fte = simulation_df["SimulationFTE"].iloc[-1]

    baseline_total_usd = simulation_df["BaselineUSD"].sum()
    simulation_total_usd = simulation_df["SimulationUSD"].sum()
    labor_cost_delta = simulation_total_usd - baseline_total_usd

    fte_delta = simulation_avg_fte - baseline_avg_fte

    with right:
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            render_custom_metric(
                "Baseline FTE",
                f"{baseline_avg_fte:,.1f}",
                "FTE at end of period",
                f"{baseline_end_fte:,.1f}",
            )

        with c2:
            render_custom_metric(
                "Simulation FTE",
                f"{simulation_avg_fte:,.1f}",
                "FTE at end of period",
                f"{simulation_end_fte:,.1f}",
                delta=f"{fte_delta:+,.1f}",
                delta_positive=fte_delta >= 0,
            )

        with c3:
            render_custom_metric(
                "Baseline Labor Cost",
                f"${baseline_total_usd:,.0f}",
            )

        with c4:
            render_custom_metric(
                "Simulation Labor Cost",
                f"${simulation_total_usd:,.0f}",
                delta=f"${labor_cost_delta:+,.0f}",
                delta_positive=labor_cost_delta >= 0,
            )

        st.markdown('<div class="metric-row-spacer"></div>', unsafe_allow_html=True)

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
                "Labor Cost",
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

        with st.expander("Audit log"):
            if st.session_state.audit_log:
                st.dataframe(pd.DataFrame(st.session_state.audit_log), use_container_width=True)
            else:
                st.info("No simulation changes applied yet.")

        with st.expander("Debug loaded monthly data"):
            st.dataframe(simulation_df, use_container_width=True)

        with st.expander("Preview loaded source files"):
            tab1, tab2, tab3 = st.tabs(
                ["FTE & Costs", "SAC Driver", "NDC Rules"]
            )

            with tab1:
                st.dataframe(st.session_state.data_sample_df.head(100))

            with tab2:
                st.dataframe(st.session_state.drivers_df.head(100))

            with tab3:
                st.dataframe(st.session_state.rules_df.head(100))


if __name__ == "__main__":
    main()
