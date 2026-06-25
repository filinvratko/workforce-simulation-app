# app.py
import os
import json
import re
import html
from datetime import datetime
from typing import Any, Dict, List, Optional

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


def generate_mock_source_data() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=36, freq="MS")
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
    months = pd.date_range("2026-01-01", periods=36, freq="MS")
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


def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
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
    missing = [col for col in required if col.lower() not in existing]

    if missing:
        raise ValueError(f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}")


def get_actual_column(df: pd.DataFrame, expected_name: str) -> Optional[str]:
    for col in df.columns:
        if str(col).strip().lower() == expected_name.lower():
            return col
    return None


def get_month_columns(df: pd.DataFrame, excluded: List[str]) -> List[str]:
    excluded_lower = [x.lower() for x in excluded]

    month_cols = [
        col
        for col in df.columns
        if str(col).strip().lower() not in excluded_lower and is_yyyymm_column(col)
    ]

    return sorted(month_cols, key=lambda x: pd.to_datetime(str(x), format="%Y%m"))


def load_data_sample(uploaded_file) -> pd.DataFrame:
    df = normalize_columns(read_excel_sheet(uploaded_file, DATA_SAMPLE_SHEET))
    validate_required_columns(df, DATA_SAMPLE_ROW_COLUMNS, DATA_SAMPLE_SHEET)

    if get_actual_column(df, "Version") is None:
        df["Version"] = "Baseline"

    month_cols = get_month_columns(df, DATA_SAMPLE_ROW_COLUMNS + ["Version"])

    if not month_cols:
        raise ValueError("No YYYYMM month columns found in Data Sample.")

    for col in month_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_drivers(uploaded_file) -> pd.DataFrame:
    df = normalize_columns(read_excel_sheet(uploaded_file, DRIVERS_SHEET))
    validate_required_columns(df, ["Driver"], DRIVERS_SHEET)
    return df


def load_calculation_rules(uploaded_file) -> pd.DataFrame:
    df = normalize_columns(read_excel_sheet(uploaded_file, CALC_RULES_SHEET))
    validate_required_columns(df, CALC_RULES_REQUIRED_COLUMNS, CALC_RULES_SHEET)
    return df


def filter_version(df: pd.DataFrame, version: str) -> pd.DataFrame:
    version_col = get_actual_column(df, "Version")

    if version_col is None:
        return df.copy()

    filtered = df[
        df[version_col].astype(str).str.strip().str.lower() == version.lower()
    ].copy()

    return filtered if not filtered.empty else df.copy()


def identify_fte_rows(df: pd.DataFrame) -> pd.Series:
    account_col = get_actual_column(df, "Account")

    if account_col is None:
        return pd.Series(False, index=df.index)

    txt = df[account_col].astype(str).str.lower().str.strip()

    return (
        txt.str.contains("fte", na=False)
        | txt.str.contains("headcount", na=False)
        | txt.str.fullmatch("hc", na=False)
    )


def source_to_monthly_baseline(source_df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_columns(source_df)
    baseline_df = filter_version(df, "Baseline")

    version_col = get_actual_column(baseline_df, "Version") or "Version"
    id_columns = DATA_SAMPLE_ROW_COLUMNS.copy()

    if version_col in baseline_df.columns:
        id_columns.append(version_col)

    month_cols = get_month_columns(baseline_df, DATA_SAMPLE_ROW_COLUMNS + ["Version"])

    long_df = baseline_df.melt(
        id_vars=id_columns,
        value_vars=month_cols,
        var_name="MonthYYYYMM",
        value_name="Value",
    )

    long_df["Month"] = pd.to_datetime(long_df["MonthYYYYMM"], format="%Y%m")
    long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0)

    fte_mask = identify_fte_rows(long_df)

    fte_df = (
        long_df[fte_mask]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "BaselineFTE"})
    )

    cost_df = (
        long_df[~fte_mask]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "BaselineUSD"})
    )

    monthly = pd.merge(fte_df, cost_df, on="Month", how="outer").fillna(0)
    monthly = monthly.sort_values("Month")
    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")

    return monthly


def extract_first_number(text: str) -> float:
    match = re.search(r"[-+]?\d*\.?\d+", text)
    return float(match.group()) if match else 0.0


def extract_month_yyyymm(text: str) -> Optional[str]:
    lower = text.lower()

    direct = re.search(r"\b(20\d{2})(0[1-9]|1[0-2])\b", lower)
    if direct:
        return direct.group()

    month_map = {
        "january": "01", "jan": "01",
        "february": "02", "feb": "02",
        "march": "03", "mar": "03",
        "april": "04", "apr": "04",
        "may": "05",
        "june": "06", "jun": "06",
        "july": "07", "jul": "07",
        "august": "08", "aug": "08",
        "september": "09", "sep": "09",
        "october": "10", "oct": "10",
        "november": "11", "nov": "11",
        "december": "12", "dec": "12",
    }

    for month_name, month_num in month_map.items():
        match_1 = re.search(rf"\b{month_name}\s+(20\d{{2}})\b", lower)
        if match_1:
            return f"{match_1.group(1)}{month_num}"

        match_2 = re.search(rf"\b(20\d{{2}})\s+{month_name}\b", lower)
        if match_2:
            return f"{match_2.group(1)}{month_num}"

    return None


def fallback_parse_instruction(text: str) -> Dict[str, Any]:
    lower = text.lower()
    month = extract_month_yyyymm(lower)

    if not month:
        return {
            "status": "clarification_required",
            "clarification_question": "What is the effective date? Please use YYYYMM, e.g. 202701.",
            "user_request": text,
            "actions": [],
            "executive_summary": "Clarification required.",
        }

    number = extract_first_number(lower)

    action = {
        "action_type": "cost_change",
        "fte_delta": 0,
        "cost_pct_delta": 0,
        "cost_abs_delta": 0,
        "effective_month": month,
        "end_month": None,
    }

    if any(w in lower for w in ["hire", "add fte", "increase fte", "recruit", "staff increase"]):
        action["action_type"] = "hire"
        action["fte_delta"] = abs(number)

    elif any(w in lower for w in ["layoff", "remove fte", "reduce fte", "cut fte", "workforce reduction"]):
        action["action_type"] = "layoff"
        action["fte_delta"] = -abs(number)

    else:
        action["action_type"] = "cost_change"
        action["cost_pct_delta"] = abs(number) / 100

        if any(w in lower for w in ["reduce", "decrease", "cut", "lower", "reduction"]):
            action["cost_pct_delta"] = -abs(action["cost_pct_delta"])

    return {
        "status": "ready_to_apply",
        "user_request": text,
        "effective_date": month,
        "changes_applied": [text],
        "actions": [action],
        "executive_summary": text[:100],
    }


def get_openai_api_key() -> Optional[str]:
    try:
        return st.secrets.get("OPENAI_API_KEY")
    except Exception:
        return os.getenv("OPENAI_API_KEY")


def call_openai_api(prompt: str) -> Dict[str, Any]:
    api_key = get_openai_api_key()

    if not api_key:
        return fallback_parse_instruction(prompt)

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)

        system_prompt = """
You are a Workforce Planning Simulation Agent.
Baseline is read-only. Simulation starts as copy of Baseline.
Effective date is mandatory.

Accept dates:
- YYYYMM
- January 2027
- Jan 2027
- 2027 January
- 2027 Jan

Return JSON only:
{
  "status": "ready_to_apply" or "clarification_required",
  "clarification_question": string or null,
  "user_request": string,
  "effective_date": "YYYYMM",
  "changes_applied": [string],
  "executive_summary": string,
  "actions": [
    {
      "action_type": "hire" | "layoff" | "cost_change",
      "fte_delta": number,
      "cost_pct_delta": number,
      "cost_abs_delta": number,
      "effective_month": "YYYYMM",
      "end_month": null
    }
  ]
}
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt},
            ],
        )

        return json.loads(response.choices[0].message.content.strip())

    except Exception:
        return fallback_parse_instruction(prompt)


def create_simulation_from_baseline(baseline_df: pd.DataFrame) -> pd.DataFrame:
    df = baseline_df.copy()
    df["SimulationFTE"] = df["BaselineFTE"]
    df["SimulationUSD"] = df["BaselineUSD"]
    return df


def apply_simulation_logic(
    baseline_monthly: pd.DataFrame,
    events: List[Dict[str, Any]],
) -> pd.DataFrame:
    df = create_simulation_from_baseline(baseline_monthly)

    avg_cost_per_fte = (
        df["BaselineUSD"].sum() / df["BaselineFTE"].sum()
        if df["BaselineFTE"].sum() else 0
    )

    for event in events:
        if event.get("status") != "ready_to_apply":
            continue

        for action in event.get("actions", []):
            effective_month = pd.to_datetime(
                str(action.get("effective_month")),
                format="%Y%m",
                errors="coerce",
            )

            if pd.isna(effective_month):
                continue

            mask = df["Month"] >= effective_month

            action_type = action.get("action_type")
            fte_delta = float(action.get("fte_delta", 0) or 0)
            cost_pct_delta = float(action.get("cost_pct_delta", 0) or 0)
            cost_abs_delta = float(action.get("cost_abs_delta", 0) or 0)

            if action_type in ["hire", "layoff"]:
                df.loc[mask, "SimulationFTE"] += fte_delta
                df.loc[mask, "SimulationUSD"] += fte_delta * avg_cost_per_fte
            else:
                df.loc[mask, "SimulationUSD"] *= 1 + cost_pct_delta
                df.loc[mask, "SimulationUSD"] += cost_abs_delta

    df["SimulationFTE"] = df["SimulationFTE"].clip(lower=0)
    df["SimulationUSD"] = df["SimulationUSD"].clip(lower=0)

    return df


def render_chart(title, df, baseline_metric, simulation_metric, y_title, height):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=df["MonthLabel"],
        y=df[baseline_metric],
        name="Baseline",
        marker_color="#2563EB",
    ))

    fig.add_trace(go.Bar(
        x=df["MonthLabel"],
        y=df[simulation_metric],
        name="Simulation",
        marker_color="#F97316",
    ))

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
            x=1,
        ),
        yaxis_title=y_title,
    )

    fig.update_yaxes(gridcolor="#374151")
    fig.update_xaxes(showgrid=False, tickangle=-45)

    return fig


def render_datapoint(
    title: str,
    main_value: str,
    secondary_label: str = "",
    secondary_value: str = "",
    delta: str = "",
    delta_positive: bool = True,
) -> None:
    title = html.escape(str(title))
    main_value = html.escape(str(main_value))
    secondary_label = html.escape(str(secondary_label))
    secondary_value = html.escape(str(secondary_value))
    delta = html.escape(str(delta))

    delta_class = "positive" if delta_positive else "negative"

    secondary_html = ""
    if secondary_label:
        secondary_html = f"""
            <div class="dp-secondary">
                {secondary_label}: <b>{secondary_value}</b>
            </div>
        """

    delta_html = ""
    if delta:
        delta_html = f"""
            <div class="dp-delta {delta_class}">
                {delta}
            </div>
        """

    st.markdown(
        f"""
        <div class="dp-card">
            <div class="dp-title">{title}</div>
            <div class="dp-main-area">
                <div class="dp-main-stack">
                    <div class="dp-main">{main_value}</div>
                    {secondary_html}
                </div>
                {delta_html}
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_summary(df: pd.DataFrame, events: List[Dict[str, Any]]) -> List[str]:
    fte_impact = df["SimulationFTE"].sum() - df["BaselineFTE"].sum()
    cost_impact = df["SimulationUSD"].sum() - df["BaselineUSD"].sum()

    return [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total labor cost impact: ${cost_impact:+,.0f}",
        "Baseline preserved: Yes",
        f"Simulation changes applied: {len(events)}",
    ]


def apply_css():
    st.markdown("""
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

    .dp-card {
        background: #111827;
        border: 1px solid #1F2937;
        border-radius: 10px;
        height: 126px;
        padding: 18px 22px;
        box-sizing: border-box;
        overflow: hidden;
    }

    .dp-title {
        color: #F9FAFB;
        font-size: 14px;
        font-weight: 700;
        line-height: 20px;
        margin-bottom: 12px;
        white-space: nowrap;
    }

    .dp-main-area {
        display: flex;
        align-items: flex-start;
        justify-content: flex-start;
        gap: 18px;
        width: 100%;
    }

    .dp-main-stack {
        display: inline-flex;
        flex-direction: column;
        align-items: center;
        justify-content: flex-start;
        width: fit-content;
        max-width: 75%;
    }

    .dp-main {
        color: #F9FAFB;
        font-size: 34px;
        line-height: 38px;
        font-weight: 800;
        white-space: nowrap;
        text-align: center;
    }

    .dp-secondary {
        color: #CBD5E1;
        font-size: 10px;
        line-height: 12px;
        margin-top: 8px;
        white-space: nowrap;
        text-align: center;
    }

    .dp-delta {
        align-self: center;
        font-size: 14px;
        line-height: 18px;
        font-weight: 800;
        white-space: nowrap;
        padding: 5px 10px;
        border-radius: 999px;
        margin-top: 3px;
    }

    .dp-delta.positive {
        color: #22C55E;
        background: rgba(34, 197, 94, 0.15);
    }

    .dp-delta.negative {
        color: #EF4444;
        background: rgba(239, 68, 68, 0.15);
    }

    .metric-row-spacer {
        height: 5px;
    }

    .box {
        background: #111827;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 14px;
        margin-bottom: 12px;
    }

    .summary {
        background: #111827;
        border: 1px solid #334155;
        border-left: 6px solid #F97316;
        border-radius: 12px;
        padding: 18px 24px;
        margin-top: 18px;
    }
    </style>
    """, unsafe_allow_html=True)


def render_header():
    st.markdown(f"""
    <div class="main-header">
        <div class="logo">TEVA</div>
        <div class="title">{REPORT_TITLE}</div>
        <div class="user">
            <div>TEST USER</div>
            <div class="avatar"></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def initialize_state():
    defaults = {
        "data_sample_df": generate_mock_source_data(),
        "drivers_df": generate_mock_drivers(),
        "rules_df": generate_mock_rules(),
        "simulation_events": [],
        "simulation_requests": [],
        "data_sample_source": "Mock data",
        "drivers_source": "Mock drivers",
        "rules_source": "Mock rules",
        "last_data_sample_upload_id": None,
        "last_drivers_upload_id": None,
        "last_rules_upload_id": None,
        "simulation_input_counter": 0,
    }

    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def main():
    initialize_state()
    apply_css()
    render_header()

    left, right = st.columns([1, 4], gap="large")

    with left:
        st.subheader("GPT Simulation")

        data_file = st.file_uploader("Upload / Reupload Data Sample", type=["xlsx"])
        if data_file is not None:
            upload_id = f"{data_file.name}_{data_file.size}"
            if upload_id != st.session_state.last_data_sample_upload_id:
                st.session_state.data_sample_df = load_data_sample(data_file)
                st.session_state.data_sample_source = data_file.name
                st.session_state.simulation_events = []
                st.session_state.simulation_requests = []
                st.session_state.last_data_sample_upload_id = upload_id
                st.rerun()

        if st.button("Reset baseline to mock data", use_container_width=True):
            st.session_state.data_sample_df = generate_mock_source_data()
            st.session_state.data_sample_source = "Mock data"
            st.session_state.simulation_events = []
            st.session_state.simulation_requests = []
            st.rerun()

        drivers_file = st.file_uploader("Upload Drivers", type=["xlsx"])
        if drivers_file is not None:
            upload_id = f"{drivers_file.name}_{drivers_file.size}"
            if upload_id != st.session_state.last_drivers_upload_id:
                st.session_state.drivers_df = load_drivers(drivers_file)
                st.session_state.drivers_source = drivers_file.name
                st.session_state.last_drivers_upload_id = upload_id
                st.rerun()

        calc_file = st.file_uploader("Upload Calculation_method", type=["xlsx"])
        if calc_file is not None:
            upload_id = f"{calc_file.name}_{calc_file.size}"
            if upload_id != st.session_state.last_rules_upload_id:
                st.session_state.rules_df = load_calculation_rules(calc_file)
                st.session_state.rules_source = calc_file.name
                st.session_state.last_rules_upload_id = upload_id
                st.rerun()

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

        input_key = f"simulation_input_text_{st.session_state.simulation_input_counter}"

        st.markdown('<div class="box">', unsafe_allow_html=True)
        simulation_prompt = st.text_area(
            "Simulation input",
            placeholder="Example: Hire 10 FTE effective January 2027",
            height=110,
            key=input_key,
        )
        run_simulation = st.button("Run simulation", use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

        if run_simulation and simulation_prompt.strip():
            prompt = simulation_prompt.strip()
            parsed = call_openai_api(prompt)

            if parsed.get("status") == "ready_to_apply":
                st.session_state.simulation_events.append(parsed)

            st.session_state.simulation_requests.insert(0, {
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "request": prompt,
                "response": parsed.get("executive_summary", ""),
                "status": parsed.get("status", ""),
            })

            st.session_state.simulation_input_counter += 1
            st.rerun()

        if st.button("Reset simulation only", use_container_width=True):
            st.session_state.simulation_events = []
            st.session_state.simulation_requests = []
            st.session_state.simulation_input_counter += 1
            st.rerun()

        st.subheader("Executed simulations")

        for item in st.session_state.simulation_requests:
            st.markdown(
                f"""
                <div class="box">
                    <b>{html.escape(item["timestamp"])}</b><br>
                    {html.escape(item["status"])}<br><br>
                    <b>Request:</b> {html.escape(item["request"])}<br>
                    <b>Result:</b> {html.escape(item["response"])}
                </div>
                """,
                unsafe_allow_html=True,
            )

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

    fte_delta = simulation_avg_fte - baseline_avg_fte
    labor_cost_delta = simulation_total_usd - baseline_total_usd

    with right:
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            render_datapoint(
                "Baseline FTE",
                f"{baseline_avg_fte:,.1f}",
                "FTE at end of period",
                f"{baseline_end_fte:,.1f}",
            )

        with c2:
            render_datapoint(
                "Simulation FTE",
                f"{simulation_avg_fte:,.1f}",
                "FTE at end of period",
                f"{simulation_end_fte:,.1f}",
                delta=f"↑ {fte_delta:+,.1f}" if fte_delta >= 0 else f"↓ {fte_delta:+,.1f}",
                delta_positive=fte_delta >= 0,
            )

        with c3:
            render_datapoint(
                "Baseline Labor Cost",
                f"${baseline_total_usd:,.0f}",
            )

        with c4:
            render_datapoint(
                "Simulation Labor Cost",
                f"${simulation_total_usd:,.0f}",
                delta=f"↑ ${labor_cost_delta:+,.0f}" if labor_cost_delta >= 0 else f"↓ ${labor_cost_delta:+,.0f}",
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
                240,
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
                480,
            ),
            use_container_width=True,
        )

        st.markdown('<div class="summary">', unsafe_allow_html=True)
        st.subheader("Simulation Impact Summary")

        for item in build_summary(simulation_df, st.session_state.simulation_events):
            st.markdown(f"- {item}")

        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("Debug loaded monthly data"):
            st.dataframe(simulation_df, use_container_width=True)

        with st.expander("Preview loaded source files"):
            tab1, tab2, tab3 = st.tabs(["FTE & Costs", "SAC Driver", "NDC Rules"])

            with tab1:
                st.dataframe(st.session_state.data_sample_df.head(100))

            with tab2:
                st.dataframe(st.session_state.drivers_df.head(100))

            with tab3:
                st.dataframe(st.session_state.rules_df.head(100))


if __name__ == "__main__":
    main()
