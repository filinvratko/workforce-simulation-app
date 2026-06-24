import os
import json
import re
from typing import Dict, List, Tuple, Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# App configuration
# ============================================================

st.set_page_config(
    page_title="Workforce Planning Simulation Report",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# Constants
# ============================================================

REPORT_TITLE = "Workforce Planning Simulation Report"
COMPANY_LOGO_URL = "https://www.cleanpng.com/png-teva-logo-transparent-png-download-2ohbhi/download-png.html"

BASELINE_REQUIRED_COLUMNS = [
    "Month",
    "Role",
    "Location",
    "FTE",
    "AverageSalary",
    "CostDriverPct",
]

DRIVERS_REQUIRED_COLUMNS = [
    "DriverName",
    "DriverType",
    "Value",
    "EffectiveMonth",
]

RULES_REQUIRED_COLUMNS = [
    "RuleName",
    "RuleType",
    "Value",
]


# ============================================================
# Session state initialization
# ============================================================

def initialize_session_state() -> None:
    if "baseline_df" not in st.session_state:
        st.session_state.baseline_df = generate_mock_data()

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
                    "Describe a workforce scenario, for example: "
                    "'Hire 10 engineers in Prague from Apr 2026' or "
                    "'Reduce salary cost by 5% from June 2026'."
                ),
            }
        ]

    if "simulation_df" not in st.session_state:
        st.session_state.simulation_df = apply_simulation_logic(
            st.session_state.baseline_df,
            st.session_state.simulation_events,
        )

    if "data_source" not in st.session_state:
        st.session_state.data_source = "Mock data"


# ============================================================
# Data generation
# ============================================================

def generate_mock_data() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=12, freq="MS")

    rows = []
    roles = ["Operator", "Engineer", "Manager", "Analyst"]
    locations = ["Prague", "Brno", "Ostrava"]

    base_values = {
        "Operator": {"fte": 80, "salary": 3200},
        "Engineer": {"fte": 45, "salary": 5200},
        "Manager": {"fte": 18, "salary": 7600},
        "Analyst": {"fte": 25, "salary": 4300},
    }

    for month_index, month in enumerate(months):
        for role in roles:
            for location in locations:
                base_fte = base_values[role]["fte"] / len(locations)
                growth = month_index * 0.15
                salary = base_values[role]["salary"]
                cost_driver_pct = 0.18

                rows.append(
                    {
                        "Month": month,
                        "Role": role,
                        "Location": location,
                        "FTE": round(base_fte + growth, 2),
                        "AverageSalary": salary,
                        "CostDriverPct": cost_driver_pct,
                    }
                )

    return pd.DataFrame(rows)


def generate_mock_drivers() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "DriverName": ["Payroll tax", "Benefits", "Annual merit increase"],
            "DriverType": ["Cost", "Cost", "Salary"],
            "Value": [0.18, 0.07, 0.03],
            "EffectiveMonth": ["2026-01-01", "2026-01-01", "2026-07-01"],
        }
    )


def generate_mock_rules() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "RuleName": ["Monthly labor cost", "FTE rounding"],
            "RuleType": ["Calculation", "Display"],
            "Value": [
                "FTE * AverageSalary * (1 + CostDriverPct)",
                "Round FTE to 2 decimals",
            ],
        }
    )


# ============================================================
# Excel loading and validation
# ============================================================

def validate_columns(
    df: pd.DataFrame,
    required_columns: List[str],
    sheet_name: str,
) -> List[str]:
    missing = [col for col in required_columns if col not in df.columns]

    if missing:
        return [f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}"]

    return []


def validate_excel_sheets(
    sheets: Dict[str, pd.DataFrame],
) -> Tuple[bool, List[str]]:
    errors = []

    required_sheets = {
        "Baseline": BASELINE_REQUIRED_COLUMNS,
        "Drivers": DRIVERS_REQUIRED_COLUMNS,
        "Rules": RULES_REQUIRED_COLUMNS,
    }

    for sheet_name, required_columns in required_sheets.items():
        if sheet_name not in sheets:
            errors.append(f"Missing required sheet: {sheet_name}")
        else:
            errors.extend(
                validate_columns(sheets[sheet_name], required_columns, sheet_name)
            )

    return len(errors) == 0, errors


def load_excel_data(uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)

    is_valid, errors = validate_excel_sheets(sheets)

    if not is_valid:
        raise ValueError("\n".join(errors))

    baseline_df = sheets["Baseline"].copy()
    drivers_df = sheets["Drivers"].copy()
    rules_df = sheets["Rules"].copy()

    baseline_df["Month"] = pd.to_datetime(baseline_df["Month"])
    baseline_df["FTE"] = pd.to_numeric(baseline_df["FTE"], errors="coerce")
    baseline_df["AverageSalary"] = pd.to_numeric(
        baseline_df["AverageSalary"],
        errors="coerce",
    )
    baseline_df["CostDriverPct"] = pd.to_numeric(
        baseline_df["CostDriverPct"],
        errors="coerce",
    )

    if baseline_df[["FTE", "AverageSalary", "CostDriverPct"]].isna().any().any():
        raise ValueError(
            "Baseline sheet contains invalid numeric values in FTE, "
            "AverageSalary, or CostDriverPct."
        )

    return baseline_df, drivers_df, rules_df


# ============================================================
# API connector placeholder
# ============================================================

def future_data_connector() -> Dict[str, Any]:
    return {
        "status": "placeholder",
        "description": (
            "Future connector for HRIS, payroll, ERP, data lake, "
            "or workforce planning system integration."
        ),
        "supported_sources": [
            "Workday",
            "SAP SuccessFactors",
            "Oracle HCM",
            "Azure SQL",
            "Snowflake",
            "SharePoint",
        ],
    }


# ============================================================
# OpenAI integration
# ============================================================

def get_openai_api_key() -> str | None:
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass

    return os.getenv("OPENAI_API_KEY")


def call_openai_api(user_instruction: str) -> Dict[str, Any]:
    api_key = get_openai_api_key()

    if not api_key:
        return fallback_parse_simulation_instruction(user_instruction)

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)

        system_prompt = """
You are a workforce planning simulation parser.

Convert the user instruction into valid JSON only.

Allowed JSON structure:
{
  "summary": "short plain English summary",
  "actions": [
    {
      "action_type": "hire | layoff | salary_change | role_change | location_change | cost_driver_change",
      "fte_delta": number,
      "salary_pct_delta": number,
      "cost_driver_pct_delta": number,
      "role": string or null,
      "location": string or null,
      "target_role": string or null,
      "target_location": string or null,
      "effective_month": "YYYY-MM-01"
    }
  ]
}

Rules:
- Hiring means positive fte_delta.
- Layoff means negative fte_delta.
- Salary reduction means negative salary_pct_delta.
- Salary increase means positive salary_pct_delta.
- Cost driver reduction means negative cost_driver_pct_delta.
- Cost driver increase means positive cost_driver_pct_delta.
- If month is missing, use 2026-01-01.
- If role or location is missing, use null.
- Return JSON only.
"""

        response = client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_instruction},
            ],
        )

        content = response.choices[0].message.content.strip()
        return json.loads(content)

    except Exception:
        return fallback_parse_simulation_instruction(user_instruction)


def fallback_parse_simulation_instruction(user_instruction: str) -> Dict[str, Any]:
    text = user_instruction.lower()

    month = extract_month(text)
    number = extract_first_number(text)

    action = {
        "action_type": "salary_change",
        "fte_delta": 0,
        "salary_pct_delta": 0,
        "cost_driver_pct_delta": 0,
        "role": extract_role(text),
        "location": extract_location(text),
        "target_role": None,
        "target_location": None,
        "effective_month": month,
    }

    if any(word in text for word in ["hire", "add", "recruit"]):
        action["action_type"] = "hire"
        action["fte_delta"] = number

    elif any(word in text for word in ["layoff", "remove", "reduce fte", "cut fte"]):
        action["action_type"] = "layoff"
        action["fte_delta"] = -abs(number)

    elif "salary" in text or "pay" in text or "wage" in text:
        action["action_type"] = "salary_change"
        action["salary_pct_delta"] = number / 100

        if any(word in text for word in ["reduce", "decrease", "cut", "lower"]):
            action["salary_pct_delta"] = -abs(action["salary_pct_delta"])

    elif "cost driver" in text or "benefit" in text or "tax" in text:
        action["action_type"] = "cost_driver_change"
        action["cost_driver_pct_delta"] = number / 100

        if any(word in text for word in ["reduce", "decrease", "cut", "lower"]):
            action["cost_driver_pct_delta"] = -abs(action["cost_driver_pct_delta"])

    elif "role" in text:
        action["action_type"] = "role_change"
        action["target_role"] = "New Role"

    elif "location" in text or "move" in text:
        action["action_type"] = "location_change"
        action["target_location"] = "New Location"

    return {
        "summary": user_instruction[:100],
        "actions": [action],
    }


def extract_first_number(text: str) -> float:
    match = re.search(r"[-+]?\d*\.?\d+", text)
    return float(match.group()) if match else 0.0


def extract_month(text: str) -> str:
    month_map = {
        "jan": "2026-01-01",
        "january": "2026-01-01",
        "feb": "2026-02-01",
        "february": "2026-02-01",
        "mar": "2026-03-01",
        "march": "2026-03-01",
        "apr": "2026-04-01",
        "april": "2026-04-01",
        "may": "2026-05-01",
        "jun": "2026-06-01",
        "june": "2026-06-01",
        "jul": "2026-07-01",
        "july": "2026-07-01",
        "aug": "2026-08-01",
        "august": "2026-08-01",
        "sep": "2026-09-01",
        "september": "2026-09-01",
        "oct": "2026-10-01",
        "october": "2026-10-01",
        "nov": "2026-11-01",
        "november": "2026-11-01",
        "dec": "2026-12-01",
        "december": "2026-12-01",
    }

    for key, value in month_map.items():
        if key in text:
            return value

    return "2026-01-01"


def extract_role(text: str) -> str | None:
    roles = ["operator", "engineer", "manager", "analyst"]

    for role in roles:
        if role in text:
            return role.title()

    return None


def extract_location(text: str) -> str | None:
    locations = ["prague", "brno", "ostrava"]

    for location in locations:
        if location in text:
            return location.title()

    return None


# ============================================================
# Simulation logic
# ============================================================

def calculate_labor_cost(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["TotalCostOfLabor"] = (
        result["FTE"]
        * result["AverageSalary"]
        * (1 + result["CostDriverPct"])
    )
    return result


def apply_simulation_logic(
    baseline_df: pd.DataFrame,
    simulation_events: List[Dict[str, Any]],
) -> pd.DataFrame:
    df = baseline_df.copy()
    df["Month"] = pd.to_datetime(df["Month"])

    for event in simulation_events:
        for action in event.get("actions", []):
            effective_month = pd.to_datetime(
                action.get("effective_month", "2026-01-01")
            )

            mask = df["Month"] >= effective_month

            role = action.get("role")
            location = action.get("location")

            if role:
                mask &= df["Role"].str.lower() == str(role).lower()

            if location:
                mask &= df["Location"].str.lower() == str(location).lower()

            action_type = action.get("action_type")

            if action_type in ["hire", "layoff"]:
                fte_delta = float(action.get("fte_delta", 0))
                month_count = max(df.loc[mask, "Month"].nunique(), 1)
                row_count = max(mask.sum(), 1)
                distributed_delta = fte_delta / row_count
                df.loc[mask, "FTE"] = df.loc[mask, "FTE"] + distributed_delta

            elif action_type == "salary_change":
                salary_pct_delta = float(action.get("salary_pct_delta", 0))
                df.loc[mask, "AverageSalary"] = (
                    df.loc[mask, "AverageSalary"] * (1 + salary_pct_delta)
                )

            elif action_type == "cost_driver_change":
                cost_driver_delta = float(action.get("cost_driver_pct_delta", 0))
                df.loc[mask, "CostDriverPct"] = (
                    df.loc[mask, "CostDriverPct"] + cost_driver_delta
                )

            elif action_type == "role_change":
                target_role = action.get("target_role")
                if target_role:
                    df.loc[mask, "Role"] = target_role

            elif action_type == "location_change":
                target_location = action.get("target_location")
                if target_location:
                    df.loc[mask, "Location"] = target_location

    df["FTE"] = df["FTE"].clip(lower=0)
    df["CostDriverPct"] = df["CostDriverPct"].clip(lower=0)

    return calculate_labor_cost(df)


def aggregate_monthly(df: pd.DataFrame) -> pd.DataFrame:
    calc_df = calculate_labor_cost(df)

    monthly = (
        calc_df.groupby("Month", as_index=False)
        .agg(
            FTE=("FTE", "sum"),
            TotalCostOfLabor=("TotalCostOfLabor", "sum"),
        )
        .sort_values("Month")
    )

    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")
    return monthly


# ============================================================
# Chart rendering
# ============================================================

def render_grouped_bar_chart(
    title: str,
    baseline_monthly: pd.DataFrame,
    simulation_monthly: pd.DataFrame,
    metric: str,
    y_axis_title: str,
) -> go.Figure:
    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=baseline_monthly["MonthLabel"],
            y=baseline_monthly[metric],
            name="Baseline",
            marker_color="#2563EB",
        )
    )

    fig.add_trace(
        go.Bar(
            x=simulation_monthly["MonthLabel"],
            y=simulation_monthly[metric],
            name="Simulation Result",
            marker_color="#F97316",
        )
    )

    fig.update_layout(
        title=title,
        barmode="group",
        height=350,
        margin=dict(l=20, r=20, t=60, b=30),
        legend=dict(orientation="h", y=1.12),
        yaxis_title=y_axis_title,
        plot_bgcolor="white",
        paper_bgcolor="white",
    )

    fig.update_yaxes(gridcolor="#E5E7EB")
    fig.update_xaxes(showgrid=False)

    return fig


# ============================================================
# Summary rendering
# ============================================================

def build_impact_summary(
    baseline_df: pd.DataFrame,
    simulation_df: pd.DataFrame,
    simulation_events: List[Dict[str, Any]],
) -> List[str]:
    baseline_monthly = aggregate_monthly(baseline_df)
    simulation_monthly = aggregate_monthly(simulation_df)

    fte_impact = simulation_monthly["FTE"].sum() - baseline_monthly["FTE"].sum()
    cost_impact = (
        simulation_monthly["TotalCostOfLabor"].sum()
        - baseline_monthly["TotalCostOfLabor"].sum()
    )

    latest_fte = simulation_monthly["FTE"].iloc[-1] - baseline_monthly["FTE"].iloc[-1]
    latest_cost = (
        simulation_monthly["TotalCostOfLabor"].iloc[-1]
        - baseline_monthly["TotalCostOfLabor"].iloc[-1]
    )

    summary = [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total labor cost impact: €{cost_impact:+,.0f}",
        f"Latest month FTE impact: {latest_fte:+,.2f}",
        f"Latest month cost impact: €{latest_cost:+,.0f}",
        f"Simulation changes applied: {len(simulation_events)}",
    ]

    if simulation_events:
        last_summary = simulation_events[-1].get("summary", "Latest change applied")
        summary.append(f"Latest change: {last_summary}")

    return [item[:100] for item in summary[:6]]


def render_impact_summary(
    baseline_df: pd.DataFrame,
    simulation_df: pd.DataFrame,
    simulation_events: List[Dict[str, Any]],
) -> None:
    st.markdown('<div class="summary-panel">', unsafe_allow_html=True)
    st.subheader("Simulation Impact Summary")

    summary = build_impact_summary(
        baseline_df,
        simulation_df,
        simulation_events,
    )

    for item in summary:
        st.markdown(f"- {item}")

    st.markdown("</div>", unsafe_allow_html=True)


# ============================================================
# UI styling
# ============================================================

def apply_custom_css() -> None:
    st.markdown(
        """
        <style>
        .block-container {
            padding-top: 1rem;
            padding-bottom: 2rem;
            max-width: 100%;
        }

        .main-header {
            width: 100%;
            background: linear-gradient(90deg, #0F172A 0%, #1E293B 100%);
            color: white;
            padding: 16px 24px;
            border-radius: 14px;
            margin-bottom: 18px;
            display: grid;
            grid-template-columns: 1fr 2fr 1fr;
            align-items: center;
            gap: 16px;
        }

        .logo-box {
            display: flex;
            align-items: center;
            gap: 12px;
            font-size: 18px;
            font-weight: 800;
        }

        .logo-placeholder {
            width: 54px;
            height: 36px;
            border-radius: 8px;
            background: white;
            color: #0F766E;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 900;
            font-size: 14px;
        }

        .report-title {
            text-align: center;
            font-size: 26px;
            font-weight: 800;
        }

        .user-box {
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 12px;
            font-size: 15px;
            font-weight: 600;
        }

        .avatar-placeholder {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            border: 2px dashed #CBD5E1;
            background: #334155;
        }

        .dashboard-panel {
            background: #FFFFFF;
            border: 1px solid #E2E8F0;
            border-radius: 16px;
            padding: 18px;
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
        }

        .chat-panel {
            background: #F8FAFC;
            border: 1px solid #E2E8F0;
            border-radius: 16px;
            padding: 16px;
            min-height: 760px;
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
        }

        .summary-panel {
            background: #FFF7ED;
            border-left: 6px solid #F97316;
            padding: 18px 24px;
            border-radius: 14px;
            margin-top: 18px;
            box-shadow: 0 8px 18px rgba(15, 23, 42, 0.05);
        }

        .small-muted {
            font-size: 13px;
            color: #64748B;
            line-height: 1.45;
        }

        div[data-testid="stMetricValue"] {
            font-size: 24px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        f"""
        <div class="main-header">
            <div class="logo-box">
                <div class="logo-placeholder">TEVA</div>
                <div>Company Logo</div>
            </div>
            <div class="report-title">{REPORT_TITLE}</div>
            <div class="user-box">
                <div>TEST USER</div>
                <div class="avatar-placeholder"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_excel_structure_help() -> None:
    with st.expander("Expected Excel structure"):
        st.markdown(
            """
            Upload one Excel file with these sheets:

            **Sheet: Baseline**

            | Month | Role | Location | FTE | AverageSalary | CostDriverPct |
            |---|---|---|---:|---:|---:|
            | 2026-01-01 | Engineer | Prague | 10 | 5200 | 0.18 |

            **Sheet: Drivers**

            | DriverName | DriverType | Value | EffectiveMonth |
            |---|---|---:|---|
            | Payroll tax | Cost | 0.18 | 2026-01-01 |

            **Sheet: Rules**

            | RuleName | RuleType | Value |
            |---|---|---|
            | Monthly labor cost | Calculation | FTE * AverageSalary * (1 + CostDriverPct) |
            """
        )


# ============================================================
# Main app
# ============================================================

def main() -> None:
    initialize_session_state()
    apply_custom_css()
    render_header()

    left_panel, right_panel = st.columns([1, 4], gap="large")

    with left_panel:
        st.markdown('<div class="chat-panel">', unsafe_allow_html=True)
        st.subheader("GPT Simulation Entry")

        uploaded_file = st.file_uploader(
            "Upload workforce Excel file",
            type=["xlsx"],
        )

        if uploaded_file:
            try:
                baseline_df, drivers_df, rules_df = load_excel_data(uploaded_file)

                st.session_state.baseline_df = baseline_df
                st.session_state.drivers_df = drivers_df
                st.session_state.rules_df = rules_df
                st.session_state.simulation_events = []
                st.session_state.simulation_df = apply_simulation_logic(
                    baseline_df,
                    [],
                )
                st.session_state.data_source = uploaded_file.name

                st.success("Excel file loaded successfully.")

            except Exception as exc:
                st.error(str(exc))

        render_excel_structure_help()

        st.caption(f"Data source: {st.session_state.data_source}")

        st.divider()

        for message in st.session_state.chat_messages:
            with st.chat_message(message["role"]):
                st.write(message["content"])

        user_prompt = st.chat_input("Enter simulation instruction...")

        if user_prompt:
            st.session_state.chat_messages.append(
                {"role": "user", "content": user_prompt}
            )

            parsed_response = call_openai_api(user_prompt)
            st.session_state.simulation_events.append(parsed_response)

            st.session_state.simulation_df = apply_simulation_logic(
                st.session_state.baseline_df,
                st.session_state.simulation_events,
            )

            assistant_response = parsed_response.get(
                "summary",
                "Simulation change applied.",
            )

            st.session_state.chat_messages.append(
                {
                    "role": "assistant",
                    "content": f"Applied: {assistant_response}",
                }
            )

            st.rerun()

        if st.button("Reset simulation", use_container_width=True):
            st.session_state.simulation_events = []
            st.session_state.simulation_df = apply_simulation_logic(
                st.session_state.baseline_df,
                [],
            )
            st.session_state.chat_messages = [
                {
                    "role": "assistant",
                    "content": "Simulation reset. Enter a new workforce scenario.",
                }
            ]
            st.rerun()

        with st.expander("API connector placeholder"):
            st.json(future_data_connector())

        st.markdown("</div>", unsafe_allow_html=True)

    with right_panel:
        baseline_monthly = aggregate_monthly(st.session_state.baseline_df)
        simulation_monthly = aggregate_monthly(st.session_state.simulation_df)

        baseline_total_fte = baseline_monthly["FTE"].sum()
        simulation_total_fte = simulation_monthly["FTE"].sum()
        baseline_total_cost = baseline_monthly["TotalCostOfLabor"].sum()
        simulation_total_cost = simulation_monthly["TotalCostOfLabor"].sum()

        st.markdown('<div class="dashboard-panel">', unsafe_allow_html=True)

        kpi_1, kpi_2, kpi_3, kpi_4 = st.columns(4)

        kpi_1.metric(
            "Baseline FTE",
            f"{baseline_total_fte:,.1f}",
        )

        kpi_2.metric(
            "Simulation FTE",
            f"{simulation_total_fte:,.1f}",
            delta=f"{simulation_total_fte - baseline_total_fte:+,.1f}",
        )

        kpi_3.metric(
            "Baseline Labor Cost",
            f"€{baseline_total_cost:,.0f}",
        )

        kpi_4.metric(
            "Simulation Labor Cost",
            f"€{simulation_total_cost:,.0f}",
            delta=f"€{simulation_total_cost - baseline_total_cost:+,.0f}",
        )

        st.plotly_chart(
            render_grouped_bar_chart(
                "FTE over Time",
                baseline_monthly,
                simulation_monthly,
                "FTE",
                "FTE",
            ),
            use_container_width=True,
        )

        st.plotly_chart(
            render_grouped_bar_chart(
                "Total Cost of Labor over Time",
                baseline_monthly,
                simulation_monthly,
                "TotalCostOfLabor",
                "Total Cost of Labor",
            ),
            use_container_width=True,
        )

        st.markdown("</div>", unsafe_allow_html=True)

        render_impact_summary(
            st.session_state.baseline_df,
            st.session_state.simulation_df,
            st.session_state.simulation_events,
        )


if __name__ == "__main__":
    main()
