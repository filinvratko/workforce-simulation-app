# app.py
import os
import json
import re
from typing import Dict, List, Any

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
    months = pd.date_range("2026-01-01", periods=24, freq="MS")
    accounts = ["FTE", "Salary", "Benefits", "Payroll Tax"]
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
                "Salary accounts use salary driver",
                "Benefits accounts use benefits driver",
                "Payroll tax accounts use tax driver",
                "FTE account identifies workforce volume",
            ],
            "Global Account": ["Salary", "Benefits", "Payroll Tax", "FTE"],
            "SAC Driver": [
                "Salary Increase",
                "Benefits Increase",
                "Payroll Tax",
                "Hiring Plan",
            ],
            "NDC Comment": [
                "Applied to salary cost",
                "Applied to benefits cost",
                "Applied to tax cost",
                "Used for FTE calculation",
            ],
            "Israel": ["Yes", "Yes", "Yes", "Yes"],
        }
    )


def read_excel_sheet(uploaded_file, sheet_name: str) -> pd.DataFrame:
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        raise ImportError(
            "Excel upload requires openpyxl. Add openpyxl to requirements.txt."
        )

    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

    if sheet_name not in sheets:
        raise ValueError(f"Missing required sheet: {sheet_name}")

    return sheets[sheet_name].copy()


def validate_required_columns(
    df: pd.DataFrame,
    required_columns: List[str],
    sheet_name: str,
) -> None:
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}"
        )


def clean_column_name(col: Any) -> str:
    text = str(col).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        text = text[:-2]
    return text


def is_yyyymm_column(column_name: Any) -> bool:
    text = clean_column_name(column_name)
    return bool(re.fullmatch(r"\d{6}", text))


def normalize_yyyymm_columns(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}

    for col in df.columns:
        text = clean_column_name(col)
        if re.fullmatch(r"\d{6}", text):
            rename_map[col] = text

    return df.rename(columns=rename_map)


def validate_yyyymm_columns(
    df: pd.DataFrame,
    excluded_columns: List[str],
    sheet: str,
) -> List[str]:
    date_columns = [
        col
        for col in df.columns
        if col not in excluded_columns and is_yyyymm_column(col)
    ]

    if not date_columns:
        raise ValueError(
            f"Sheet '{sheet}' must contain monthly date columns in YYYYMM format, e.g. 202601."
        )

    return sorted(date_columns, key=lambda x: pd.to_datetime(str(x), format="%Y%m"))


def load_data_sample(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, DATA_SAMPLE_SHEET)
    df = normalize_yyyymm_columns(df)

    validate_required_columns(df, DATA_SAMPLE_ROW_COLUMNS, DATA_SAMPLE_SHEET)

    if "Version" not in df.columns:
        df["Version"] = "Baseline"

    month_columns = validate_yyyymm_columns(
        df,
        DATA_SAMPLE_ROW_COLUMNS + ["Version"],
        DATA_SAMPLE_SHEET,
    )

    for col in month_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_drivers(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, DRIVERS_SHEET)
    df = normalize_yyyymm_columns(df)

    validate_required_columns(df, ["Driver"], DRIVERS_SHEET)

    month_columns = validate_yyyymm_columns(df, ["Driver"], DRIVERS_SHEET)

    for col in month_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def load_calculation_rules(uploaded_file) -> pd.DataFrame:
    df = read_excel_sheet(uploaded_file, CALC_RULES_SHEET)

    validate_required_columns(
        df,
        CALC_RULES_REQUIRED_COLUMNS,
        CALC_RULES_SHEET,
    )

    return df


def get_month_columns(df: pd.DataFrame, excluded_columns: List[str]) -> List[str]:
    month_columns = [
        col
        for col in df.columns
        if col not in excluded_columns and is_yyyymm_column(col)
    ]

    return sorted(month_columns, key=lambda x: pd.to_datetime(str(x), format="%Y%m"))


def source_to_monthly_baseline(source_df: pd.DataFrame) -> pd.DataFrame:
    df = normalize_yyyymm_columns(source_df)

    month_columns = get_month_columns(
        df,
        DATA_SAMPLE_ROW_COLUMNS + ["Version"],
    )

    long_df = df.melt(
        id_vars=DATA_SAMPLE_ROW_COLUMNS + ["Version"],
        value_vars=month_columns,
        var_name="MonthYYYYMM",
        value_name="Value",
    )

    long_df["Month"] = pd.to_datetime(long_df["MonthYYYYMM"], format="%Y%m")
    long_df["Value"] = pd.to_numeric(long_df["Value"], errors="coerce").fillna(0)

    fte_df = (
        long_df[
            long_df["Account"]
            .astype(str)
            .str.lower()
            .str.contains("fte", na=False)
        ]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "FTE"})
    )

    cost_df = (
        long_df[
            ~long_df["Account"]
            .astype(str)
            .str.lower()
            .str.contains("fte", na=False)
        ]
        .groupby("Month", as_index=False)["Value"]
        .sum()
        .rename(columns={"Value": "TotalCostOfLabor"})
    )

    monthly = pd.merge(fte_df, cost_df, on="Month", how="outer").fillna(0)
    monthly = monthly.sort_values("Month")
    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")

    return monthly


def get_openai_api_key():
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
You are a workforce planning simulation parser.
Return valid JSON only.

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

Rules:
- Hire means positive fte_delta.
- Layoff means negative fte_delta.
- Salary/cost increase means positive cost_pct_delta.
- Salary/cost decrease means negative cost_pct_delta.
- Use 202601 if no month is provided.
- Percentages must be decimals, e.g. 5% = 0.05.
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

    elif any(w in text for w in ["salary", "cost", "labor cost", "wage"]):
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


def apply_simulation_logic(
    baseline_monthly: pd.DataFrame,
    events: List[Dict[str, Any]],
) -> pd.DataFrame:
    df = baseline_monthly.copy()

    df["SimulationFTE"] = df["FTE"]
    df["SimulationCost"] = df["TotalCostOfLabor"]

    average_cost_per_fte = (
        df["TotalCostOfLabor"].sum() / df["FTE"].sum()
        if df["FTE"].sum() != 0
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
                df.loc[mask, "SimulationCost"] += fte_delta * average_cost_per_fte

            elif action_type in ["salary_change", "cost_change", "driver_change"]:
                df.loc[mask, "SimulationCost"] *= 1 + cost_pct_delta
                df.loc[mask, "SimulationCost"] += cost_abs_delta

    df["SimulationFTE"] = df["SimulationFTE"].clip(lower=0)
    df["SimulationCost"] = df["SimulationCost"].clip(lower=0)

    return df


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
            name="Simulation Result",
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
        margin=dict(l=30, r=30, t=55, b=85),
        legend=dict(orientation="h", y=1.14),
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
    fte_impact = df["SimulationFTE"].sum() - df["FTE"].sum()
    cost_impact = df["SimulationCost"].sum() - df["TotalCostOfLabor"].sum()

    summary = [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total labor cost impact: €{cost_impact:+,.0f}",
        f"Changes applied: {len(events)}",
    ]

    if events:
        summary.append(f"Latest change: {events[-1].get('summary', '')}")

    return [item[:100] for item in summary[:6]]


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
            padding: 16px;
            border-radius: 10px;
            border: 1px solid #1F2937;
            min-height: 118px;
            height: 118px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        div[data-testid="stMetricLabel"],
        div[data-testid="stMetricValue"] {
            color: #F9FAFB;
        }

        div[data-testid="stMetricValue"] {
            font-size: 28px;
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
                "content": "Enter scenario, e.g. hire 10 FTE from 202604 or reduce cost by 5%.",
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
    simulation_df = apply_simulation_logic(
        baseline_monthly,
        st.session_state.simulation_events,
    )

    baseline_total_fte = simulation_df["FTE"].sum()
    simulation_total_fte = simulation_df["SimulationFTE"].sum()
    baseline_total_cost = simulation_df["TotalCostOfLabor"].sum()
    simulation_total_cost = simulation_df["SimulationCost"].sum()

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
                    "content": f"Applied: {parsed.get('summary', 'Simulation change')}",
                }
            )

            st.rerun()

        if st.button("Reset simulation", use_container_width=True):
            st.session_state.simulation_events = []
            st.session_state.chat_messages = [
                {
                    "role": "assistant",
                    "content": "Simulation reset. Enter a new scenario.",
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
        c3.metric("Baseline Labor Cost", f"€{baseline_total_cost:,.0f}")
        c4.metric(
            "Simulation Labor Cost",
            f"€{simulation_total_cost:,.0f}",
            delta=f"€{simulation_total_cost - baseline_total_cost:+,.0f}",
        )

        st.plotly_chart(
            render_chart(
                "FTE over Time",
                simulation_df,
                "FTE",
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
                "TotalCostOfLabor",
                "SimulationCost",
                "Total Cost of Labor",
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
            tab1, tab2, tab3 = st.tabs(["FTE & Costs", "SAC Driver", "NDC Rules"])

            with tab1:
                st.dataframe(st.session_state.data_sample_df.head(50))

            with tab2:
                st.dataframe(st.session_state.drivers_df.head(50))

            with tab3:
                st.dataframe(st.session_state.rules_df.head(50))


if __name__ == "__main__":
    main()
    
