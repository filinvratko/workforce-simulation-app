import os
import json
import re
from typing import Dict, List, Tuple, Any

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Workforce Planning Simulation Report",
    layout="wide",
    initial_sidebar_state="collapsed",
)

REPORT_TITLE = "Workforce Planning Simulation Report"

BASELINE_REQUIRED_COLUMNS = [
    "Month", "Role", "Location", "FTE", "AverageSalary", "CostDriverPct"
]

DRIVERS_REQUIRED_COLUMNS = [
    "DriverName", "DriverType", "Value", "EffectiveMonth"
]

RULES_REQUIRED_COLUMNS = [
    "RuleName", "RuleType", "Value"
]


def generate_mock_data() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=12, freq="MS")
    roles = ["Operator", "Engineer", "Manager", "Analyst"]
    locations = ["Prague", "Brno", "Ostrava"]

    base = {
        "Operator": {"fte": 80, "salary": 3200},
        "Engineer": {"fte": 45, "salary": 5200},
        "Manager": {"fte": 18, "salary": 7600},
        "Analyst": {"fte": 25, "salary": 4300},
    }

    rows = []
    for i, month in enumerate(months):
        for role in roles:
            for location in locations:
                rows.append({
                    "Month": month,
                    "Role": role,
                    "Location": location,
                    "FTE": round(base[role]["fte"] / 3 + i * 0.15, 2),
                    "AverageSalary": base[role]["salary"],
                    "CostDriverPct": 0.18,
                })

    return pd.DataFrame(rows)


def generate_mock_drivers() -> pd.DataFrame:
    return pd.DataFrame({
        "DriverName": ["Payroll tax", "Benefits", "Annual merit increase"],
        "DriverType": ["Cost", "Cost", "Salary"],
        "Value": [0.18, 0.07, 0.03],
        "EffectiveMonth": ["2026-01-01", "2026-01-01", "2026-07-01"],
    })


def generate_mock_rules() -> pd.DataFrame:
    return pd.DataFrame({
        "RuleName": ["Monthly labor cost", "FTE rounding"],
        "RuleType": ["Calculation", "Display"],
        "Value": [
            "FTE * AverageSalary * (1 + CostDriverPct)",
            "Round FTE to 2 decimals",
        ],
    })


def validate_columns(df: pd.DataFrame, required: List[str], sheet: str) -> List[str]:
    missing = [col for col in required if col not in df.columns]
    return [f"Sheet '{sheet}' is missing: {', '.join(missing)}"] if missing else []


def validate_excel_sheets(sheets: Dict[str, pd.DataFrame]) -> Tuple[bool, List[str]]:
    errors = []

    required_sheets = {
        "Baseline": BASELINE_REQUIRED_COLUMNS,
        "Drivers": DRIVERS_REQUIRED_COLUMNS,
        "Rules": RULES_REQUIRED_COLUMNS,
    }

    for sheet, columns in required_sheets.items():
        if sheet not in sheets:
            errors.append(f"Missing required sheet: {sheet}")
        else:
            errors.extend(validate_columns(sheets[sheet], columns, sheet))

    return len(errors) == 0, errors


def load_excel_data(uploaded_file):
    try:
        import openpyxl  # noqa: F401
    except ImportError:
        raise ImportError(
            "Excel upload requires openpyxl. Add 'openpyxl' to requirements.txt and redeploy."
        )

    sheets = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

    valid, errors = validate_excel_sheets(sheets)
    if not valid:
        raise ValueError("\n".join(errors))

    baseline = sheets["Baseline"].copy()
    drivers = sheets["Drivers"].copy()
    rules = sheets["Rules"].copy()

    baseline["Month"] = pd.to_datetime(baseline["Month"])
    baseline["FTE"] = pd.to_numeric(baseline["FTE"], errors="coerce")
    baseline["AverageSalary"] = pd.to_numeric(baseline["AverageSalary"], errors="coerce")
    baseline["CostDriverPct"] = pd.to_numeric(baseline["CostDriverPct"], errors="coerce")

    if baseline[["FTE", "AverageSalary", "CostDriverPct"]].isna().any().any():
        raise ValueError("Baseline has invalid numeric values.")

    return baseline, drivers, rules


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
        return fallback_parse_simulation_instruction(user_instruction)

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)

        system_prompt = """
Return valid JSON only.

Format:
{
  "summary": "short summary",
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

Use 2026-01-01 if no month is provided.
Use decimal percentages, so 5% = 0.05.
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
        return fallback_parse_simulation_instruction(user_instruction)


def fallback_parse_simulation_instruction(text: str) -> Dict[str, Any]:
    original = text
    text = text.lower()
    number = extract_first_number(text)
    month = extract_month(text)

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

    if any(w in text for w in ["hire", "add", "recruit"]):
        action["action_type"] = "hire"
        action["fte_delta"] = number

    elif any(w in text for w in ["layoff", "remove", "reduce fte", "cut fte"]):
        action["action_type"] = "layoff"
        action["fte_delta"] = -abs(number)

    elif any(w in text for w in ["salary", "pay", "wage"]):
        action["action_type"] = "salary_change"
        action["salary_pct_delta"] = number / 100
        if any(w in text for w in ["reduce", "decrease", "cut", "lower"]):
            action["salary_pct_delta"] = -abs(action["salary_pct_delta"])

    elif any(w in text for w in ["cost driver", "benefit", "tax"]):
        action["action_type"] = "cost_driver_change"
        action["cost_driver_pct_delta"] = number / 100
        if any(w in text for w in ["reduce", "decrease", "cut", "lower"]):
            action["cost_driver_pct_delta"] = -abs(action["cost_driver_pct_delta"])

    return {"summary": original[:100], "actions": [action]}


def extract_first_number(text: str) -> float:
    match = re.search(r"[-+]?\d*\.?\d+", text)
    return float(match.group()) if match else 0.0


def extract_month(text: str) -> str:
    month_map = {
        "jan": "2026-01-01", "january": "2026-01-01",
        "feb": "2026-02-01", "february": "2026-02-01",
        "mar": "2026-03-01", "march": "2026-03-01",
        "apr": "2026-04-01", "april": "2026-04-01",
        "may": "2026-05-01",
        "jun": "2026-06-01", "june": "2026-06-01",
        "jul": "2026-07-01", "july": "2026-07-01",
        "aug": "2026-08-01", "august": "2026-08-01",
        "sep": "2026-09-01", "september": "2026-09-01",
        "oct": "2026-10-01", "october": "2026-10-01",
        "nov": "2026-11-01", "november": "2026-11-01",
        "dec": "2026-12-01", "december": "2026-12-01",
    }

    for key, value in month_map.items():
        if key in text:
            return value

    return "2026-01-01"


def extract_role(text: str):
    for role in ["operator", "engineer", "manager", "analyst"]:
        if role in text:
            return role.title()
    return None


def extract_location(text: str):
    for location in ["prague", "brno", "ostrava"]:
        if location in text:
            return location.title()
    return None


def calculate_labor_cost(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    result["TotalCostOfLabor"] = (
        result["FTE"] * result["AverageSalary"] * (1 + result["CostDriverPct"])
    )
    return result


def apply_simulation_logic(baseline_df, simulation_events):
    df = baseline_df.copy()
    df["Month"] = pd.to_datetime(df["Month"])

    for event in simulation_events:
        for action in event.get("actions", []):
            month = pd.to_datetime(action.get("effective_month", "2026-01-01"))
            mask = df["Month"] >= month

            role = action.get("role")
            location = action.get("location")

            if role:
                mask &= df["Role"].str.lower() == str(role).lower()

            if location:
                mask &= df["Location"].str.lower() == str(location).lower()

            action_type = action.get("action_type")

            if action_type in ["hire", "layoff"]:
                delta = float(action.get("fte_delta", 0))
                rows = max(mask.sum(), 1)
                df.loc[mask, "FTE"] += delta / rows

            elif action_type == "salary_change":
                delta = float(action.get("salary_pct_delta", 0))
                df.loc[mask, "AverageSalary"] *= 1 + delta

            elif action_type == "cost_driver_change":
                delta = float(action.get("cost_driver_pct_delta", 0))
                df.loc[mask, "CostDriverPct"] += delta

    df["FTE"] = df["FTE"].clip(lower=0)
    df["CostDriverPct"] = df["CostDriverPct"].clip(lower=0)

    return calculate_labor_cost(df)


def aggregate_monthly(df):
    df = calculate_labor_cost(df)

    monthly = (
        df.groupby("Month", as_index=False)
        .agg(
            FTE=("FTE", "sum"),
            TotalCostOfLabor=("TotalCostOfLabor", "sum"),
        )
        .sort_values("Month")
    )

    monthly["MonthLabel"] = monthly["Month"].dt.strftime("%b %Y")
    return monthly


def render_chart(title, baseline, simulation, metric, y_title):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=baseline["MonthLabel"],
        y=baseline[metric],
        name="Baseline",
        marker_color="#2563EB",
    ))

    fig.add_trace(go.Bar(
        x=simulation["MonthLabel"],
        y=simulation[metric],
        name="Simulation Result",
        marker_color="#F97316",
    ))

    fig.update_layout(
        title=title,
        barmode="group",
        height=360,
        plot_bgcolor="#111827",
        paper_bgcolor="#111827",
        font=dict(color="#F9FAFB"),
        margin=dict(l=30, r=30, t=60, b=30),
        legend=dict(orientation="h", y=1.1),
        yaxis_title=y_title,
    )

    fig.update_yaxes(gridcolor="#374151")
    fig.update_xaxes(showgrid=False)

    return fig


def build_summary(baseline_df, simulation_df, events):
    baseline = aggregate_monthly(baseline_df)
    simulation = aggregate_monthly(simulation_df)

    fte_impact = simulation["FTE"].sum() - baseline["FTE"].sum()
    cost_impact = simulation["TotalCostOfLabor"].sum() - baseline["TotalCostOfLabor"].sum()

    summary = [
        f"Total FTE impact: {fte_impact:+,.2f}",
        f"Total labor cost impact: €{cost_impact:+,.0f}",
        f"Changes applied: {len(events)}",
    ]

    if events:
        summary.append(f"Latest change: {events[-1].get('summary', '')}")

    return [x[:100] for x in summary[:6]]


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

    div[data-testid="stMetric"] {
        background: #111827;
        padding: 16px;
        border-radius: 10px;
        border: 1px solid #1F2937;
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


def render_excel_help():
    with st.expander("Expected Excel structure"):
        st.markdown("""
        Required sheets:

        **Baseline**
        - Month
        - Role
        - Location
        - FTE
        - AverageSalary
        - CostDriverPct

        **Drivers**
        - DriverName
        - DriverType
        - Value
        - EffectiveMonth

        **Rules**
        - RuleName
        - RuleType
        - Value
        """)


def initialize_state():
    if "baseline_df" not in st.session_state:
        st.session_state.baseline_df = generate_mock_data()

    if "drivers_df" not in st.session_state:
        st.session_state.drivers_df = generate_mock_drivers()

    if "rules_df" not in st.session_state:
        st.session_state.rules_df = generate_mock_rules()

    if "simulation_events" not in st.session_state:
        st.session_state.simulation_events = []

    if "simulation_df" not in st.session_state:
        st.session_state.simulation_df = apply_simulation_logic(
            st.session_state.baseline_df,
            st.session_state.simulation_events,
        )

    if "chat_messages" not in st.session_state:
        st.session_state.chat_messages = [
            {
                "role": "assistant",
                "content": "Enter a simulation instruction, e.g. hire 10 engineers from April.",
            }
        ]

    if "data_source" not in st.session_state:
        st.session_state.data_source = "Mock data"


def main():
    initialize_state()
    apply_css()
    render_header()

    baseline_monthly = aggregate_monthly(st.session_state.baseline_df)
    simulation_monthly = aggregate_monthly(st.session_state.simulation_df)

    baseline_total_fte = baseline_monthly["FTE"].sum()
    simulation_total_fte = simulation_monthly["FTE"].sum()
    baseline_total_cost = baseline_monthly["TotalCostOfLabor"].sum()
    simulation_total_cost = simulation_monthly["TotalCostOfLabor"].sum()

    left, right = st.columns([1, 4], gap="large")

    with left:
        st.subheader("GPT Simulation")

        uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

        if uploaded_file:
            try:
                baseline, drivers, rules = load_excel_data(uploaded_file)

                st.session_state.baseline_df = baseline
                st.session_state.drivers_df = drivers
                st.session_state.rules_df = rules
                st.session_state.simulation_events = []
                st.session_state.simulation_df = apply_simulation_logic(baseline, [])
                st.session_state.data_source = uploaded_file.name

                st.success("Excel loaded successfully.")
                st.rerun()

            except ImportError as exc:
                st.warning(str(exc))

            except Exception as exc:
                st.warning(f"Excel file could not be loaded: {exc}")

        render_excel_help()
        st.caption(f"Data source: {st.session_state.data_source}")
        st.divider()

        for msg in st.session_state.chat_messages:
            with st.chat_message(msg["role"]):
                st.write(msg["content"])

        prompt = st.chat_input("Enter simulation instruction...")

        if prompt:
            st.session_state.chat_messages.append({"role": "user", "content": prompt})

            parsed = call_openai_api(prompt)
            st.session_state.simulation_events.append(parsed)
            st.session_state.simulation_df = apply_simulation_logic(
                st.session_state.baseline_df,
                st.session_state.simulation_events,
            )

            st.session_state.chat_messages.append({
                "role": "assistant",
                "content": f"Applied: {parsed.get('summary', 'Simulation change')}",
            })

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
                baseline_monthly,
                simulation_monthly,
                "FTE",
                "FTE",
            ),
            use_container_width=True,
        )

        st.plotly_chart(
            render_chart(
                "Total Cost of Labor over Time",
                baseline_monthly,
                simulation_monthly,
                "TotalCostOfLabor",
                "Total Cost of Labor",
            ),
            use_container_width=True,
        )

        st.markdown('<div class="summary">', unsafe_allow_html=True)
        st.subheader("Simulation Impact Summary")

        for item in build_summary(
            st.session_state.baseline_df,
            st.session_state.simulation_df,
            st.session_state.simulation_events,
        ):
            st.markdown(f"- {item}")

        st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
