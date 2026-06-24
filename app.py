import json
import os
import re
from typing import Any, Dict, List, Tuple

import pandas as pd
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="Workforce Planning Simulation Report",
    page_icon="W",
    layout="wide",
    initial_sidebar_state="collapsed",
)

LOGO_URL = "https://www.cleanpng.com/png-teva-logo-transparent-png-download-2ohbhi/download-png.html"

BASELINE_SHEET = "Baseline"
DRIVERS_SHEET = "Drivers"
RULES_SHEET = "Rules"

BASELINE_COLUMNS = {
    "month",
    "fte",
    "total_cost_of_labor",
    "role",
    "location",
    "average_salary",
}
DRIVERS_COLUMNS = {"driver", "value", "description"}
RULES_COLUMNS = {"rule", "value", "description"}

SIMULATION_SCHEMA = {
    "actions": [
        {
            "type": "hire | layoff | salary_change | role_change | location_change | cost_driver_change",
            "month": "YYYY-MM",
            "fte_delta": 0,
            "cost_delta": 0,
            "salary_pct_change": 0,
            "role": "",
            "location": "",
            "description": "",
        }
    ]
}


def inject_css() -> None:
    st.markdown(
        """
        <style>
            .block-container {
                padding-top: 0.75rem;
                padding-left: 1.25rem;
                padding-right: 1.25rem;
                max-width: 100%;
            }
            .app-header {
                width: 100%;
                display: grid;
                grid-template-columns: 220px 1fr 240px;
                align-items: center;
                padding: 14px 20px;
                border-bottom: 1px solid #d9dee8;
                background: #ffffff;
                margin-bottom: 16px;
            }
            .header-logo {
                display: flex;
                align-items: center;
                gap: 10px;
                font-weight: 700;
                color: #14345b;
                font-size: 15px;
            }
            .header-logo img {
                height: 42px;
                max-width: 145px;
                object-fit: contain;
            }
            .header-title {
                text-align: center;
                font-size: 25px;
                font-weight: 750;
                color: #14233b;
            }
            .header-user {
                display: flex;
                justify-content: flex-end;
                align-items: center;
                gap: 12px;
                color: #27364f;
                font-size: 14px;
                font-weight: 650;
            }
            .avatar-placeholder {
                width: 38px;
                height: 38px;
                border-radius: 50%;
                border: 1px solid #b9c2d1;
                background: #f5f7fb;
            }
            .left-panel {
                min-height: calc(100vh - 120px);
                padding: 16px;
                border: 1px solid #dbe1ea;
                background: #f8fafd;
                border-radius: 8px;
            }
            .panel-title {
                font-size: 17px;
                font-weight: 750;
                color: #172033;
                margin-bottom: 6px;
            }
            .panel-subtitle {
                font-size: 12px;
                color: #5d6b82;
                line-height: 1.35;
                margin-bottom: 14px;
            }
            .metric-row {
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 12px;
                margin-bottom: 14px;
            }
            .metric-card {
                border: 1px solid #dbe1ea;
                background: #ffffff;
                border-radius: 8px;
                padding: 12px 14px;
            }
            .metric-label {
                font-size: 12px;
                color: #63708a;
                margin-bottom: 4px;
            }
            .metric-value {
                font-size: 20px;
                font-weight: 760;
                color: #14233b;
            }
            .summary-box {
                border: 1px solid #dbe1ea;
                background: #ffffff;
                border-radius: 8px;
                padding: 16px 18px;
                margin-top: 14px;
            }
            .summary-title {
                font-size: 16px;
                font-weight: 750;
                color: #172033;
                margin-bottom: 8px;
            }
            .api-placeholder {
                border: 1px dashed #aeb8c8;
                background: #fbfcfe;
                border-radius: 8px;
                padding: 12px;
                color: #56647a;
                font-size: 12px;
                line-height: 1.35;
                margin-top: 12px;
            }
            .stChatMessage {
                background: #ffffff;
                border: 1px solid #e2e6ee;
                border-radius: 8px;
                padding: 4px;
            }
            div[data-testid="stFileUploader"] {
                background: #ffffff;
                border: 1px solid #e2e6ee;
                border-radius: 8px;
                padding: 8px;
            }
            @media (max-width: 900px) {
                .app-header {
                    grid-template-columns: 1fr;
                    gap: 10px;
                    text-align: center;
                }
                .header-user,
                .header-logo {
                    justify-content: center;
                }
                .metric-row {
                    grid-template-columns: 1fr;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def initialize_session_state() -> None:
    if "baseline_df" not in st.session_state:
        st.session_state.baseline_df = generate_mock_data()
    if "drivers_df" not in st.session_state:
        st.session_state.drivers_df = pd.DataFrame(
            [
                {
                    "driver": "benefits_load_pct",
                    "value": 0.24,
                    "description": "Benefits and payroll burden as percentage of salary",
                },
                {
                    "driver": "annual_merit_pct",
                    "value": 0.03,
                    "description": "Annual merit increase assumption",
                },
            ]
        )
    if "rules_df" not in st.session_state:
        st.session_state.rules_df = pd.DataFrame(
            [
                {
                    "rule": "cost_formula",
                    "value": "fte * average_salary / 12 * (1 + benefits_load_pct)",
                    "description": "Monthly total labor cost calculation",
                },
                {
                    "rule": "fte_precision",
                    "value": "2",
                    "description": "FTE values rounded to two decimals",
                },
            ]
        )
    if "simulation_df" not in st.session_state:
        st.session_state.simulation_df = st.session_state.baseline_df.copy()
    if "chat_history" not in st.session_state:
        st.session_state.chat_history = [
            {
                "role": "assistant",
                "content": "Enter a workforce simulation, such as: Hire 10 FTE in Prague starting 2026-04 at 72000 salary.",
            }
        ]
    if "summary" not in st.session_state:
        st.session_state.summary = [
            "No simulation changes applied yet.",
            "FTE impact: 0.00",
            "Labor cost impact: $0",
        ]
    if "validation_messages" not in st.session_state:
        st.session_state.validation_messages = []
    if "applied_actions" not in st.session_state:
        st.session_state.applied_actions = []


def generate_mock_data() -> pd.DataFrame:
    months = pd.date_range("2026-01-01", periods=12, freq="MS")
    records = []
    role_mix = [
        ("Commercial", "Prague", 82, 72000),
        ("Operations", "Dublin", 105, 68000),
        ("Manufacturing", "Berlin", 128, 64000),
        ("Finance", "Prague", 36, 78000),
        ("Technology", "Tel Aviv", 54, 92000),
    ]

    for month_index, month in enumerate(months):
        for role, location, base_fte, salary in role_mix:
            growth = month_index * 0.35
            seasonal = 1.5 if month.month in [6, 7, 8] and role == "Manufacturing" else 0
            fte = round(base_fte + growth + seasonal, 2)
            average_salary = salary * (1 + 0.002 * month_index)
            total_cost = fte * average_salary / 12 * 1.24
            records.append(
                {
                    "month": month.strftime("%Y-%m"),
                    "role": role,
                    "location": location,
                    "fte": fte,
                    "average_salary": round(average_salary, 2),
                    "total_cost_of_labor": round(total_cost, 2),
                }
            )

    return pd.DataFrame(records)


def validate_excel_sheets(excel_file: Any) -> Tuple[bool, List[str]]:
    messages = []
    try:
        workbook = pd.ExcelFile(excel_file)
    except Exception as exc:
        return False, [f"Unable to read Excel file: {exc}"]

    required_sheets = {BASELINE_SHEET, DRIVERS_SHEET, RULES_SHEET}
    missing_sheets = required_sheets - set(workbook.sheet_names)
    if missing_sheets:
        messages.append("Missing required sheet(s): " + ", ".join(sorted(missing_sheets)))

    requirements = {
        BASELINE_SHEET: BASELINE_COLUMNS,
        DRIVERS_SHEET: DRIVERS_COLUMNS,
        RULES_SHEET: RULES_COLUMNS,
    }
    for sheet, required_columns in requirements.items():
        if sheet not in workbook.sheet_names:
            continue
        try:
            df = pd.read_excel(workbook, sheet_name=sheet)
        except Exception as exc:
            messages.append(f"Unable to read sheet '{sheet}': {exc}")
            continue
        normalized_columns = {str(col).strip().lower().replace(" ", "_") for col in df.columns}
        missing_columns = required_columns - normalized_columns
        if missing_columns:
            messages.append(f"Sheet '{sheet}' missing column(s): " + ", ".join(sorted(missing_columns)))

    return len(messages) == 0, messages


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    clean_df = df.copy()
    clean_df.columns = [str(col).strip().lower().replace(" ", "_") for col in clean_df.columns]
    return clean_df


def load_excel_data(excel_file: Any) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    baseline_df = normalize_columns(pd.read_excel(excel_file, sheet_name=BASELINE_SHEET))
    drivers_df = normalize_columns(pd.read_excel(excel_file, sheet_name=DRIVERS_SHEET))
    rules_df = normalize_columns(pd.read_excel(excel_file, sheet_name=RULES_SHEET))

    baseline_df["month"] = baseline_df["month"].astype(str).str.slice(0, 7)
    baseline_df["fte"] = pd.to_numeric(baseline_df["fte"], errors="coerce").fillna(0)
    baseline_df["average_salary"] = pd.to_numeric(baseline_df["average_salary"], errors="coerce").fillna(0)
    baseline_df["total_cost_of_labor"] = pd.to_numeric(
        baseline_df["total_cost_of_labor"], errors="coerce"
    ).fillna(0)

    return baseline_df, drivers_df, rules_df


def expected_excel_structure() -> None:
    with st.expander("Expected Excel sheet structure", expanded=False):
        st.markdown(
            """
            Upload an `.xlsx` workbook with exactly these sheets:

            **Baseline**
            - `month`: Month in `YYYY-MM` format
            - `role`: Workforce role or job family
            - `location`: Workforce location
            - `fte`: Full-time equivalent headcount
            - `average_salary`: Annual average salary
            - `total_cost_of_labor`: Monthly total labor cost

            **Drivers**
            - `driver`: Driver name, such as `benefits_load_pct`
            - `value`: Driver value
            - `description`: Business description

            **Rules**
            - `rule`: Calculation rule name
            - `value`: Rule expression or parameter
            - `description`: Business description
            """
        )


def get_openai_api_key() -> str:
    try:
        if "OPENAI_API_KEY" in st.secrets:
            return st.secrets["OPENAI_API_KEY"]
    except Exception:
        pass
    return os.getenv("OPENAI_API_KEY", "")


def call_openai_api(user_instruction: str, context: Dict[str, Any]) -> Dict[str, Any]:
    api_key = get_openai_api_key()
    if not api_key:
        return parse_simulation_instruction_fallback(user_instruction)

    system_prompt = f"""
You are a workforce planning simulation parser.

Return only valid JSON matching this schema:
{json.dumps(SIMULATION_SCHEMA, indent=2)}

Rules:
- Interpret hiring as positive fte_delta.
- Interpret layoffs as negative fte_delta.
- Interpret salary changes as salary_pct_change decimal, for example 5% = 0.05.
- If a date is missing, use the earliest available month from context.
- If cost impact is explicit, place it in cost_delta.
- Keep descriptions concise and business friendly.
- Do not include markdown or explanatory text.
"""
    user_prompt = {
        "instruction": user_instruction,
        "available_months": context.get("months", []),
        "roles": context.get("roles", []),
        "locations": context.get("locations", []),
        "drivers": context.get("drivers", []),
        "rules": context.get("rules", []),
    }

    try:
        from openai import OpenAI

        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model=os.getenv("OPENAI_MODEL", "gpt-4o-mini"),
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(user_prompt)},
            ],
            temperature=0,
            response_format={"type": "json_object"},
        )
        content = response.choices[0].message.content
        parsed = json.loads(content)
        return normalize_simulation_response(parsed, user_instruction)
    except Exception:
        return parse_simulation_instruction_fallback(user_instruction)


def parse_simulation_instruction_fallback(user_instruction: str) -> Dict[str, Any]:
    text = user_instruction.lower()
    action_type = "cost_driver_change"
    fte_delta = 0.0
    cost_delta = 0.0
    salary_pct_change = 0.0

    month_match = re.search(r"(20\d{2})[-/](0[1-9]|1[0-2])", text)
    month = month_match.group(0).replace("/", "-") if month_match else ""
    number_match = re.search(r"(-?\d+(?:\.\d+)?)", text)
    amount = float(number_match.group(1)) if number_match else 0.0
    pct_match = re.search(r"(-?\d+(?:\.\d+)?)\s?%", text)
    money_match = re.search(r"\$?\s?(-?\d+(?:,\d{3})*(?:\.\d+)?)", text)

    if any(word in text for word in ["hire", "hiring", "add", "recruit"]):
        action_type = "hire"
        fte_delta = abs(amount)
    elif any(word in text for word in ["layoff", "lay off", "reduce", "reduction", "terminate"]):
        action_type = "layoff"
        fte_delta = -abs(amount)
    elif any(word in text for word in ["salary", "merit", "pay", "wage", "compensation"]):
        action_type = "salary_change"
        salary_pct_change = float(pct_match.group(1)) / 100 if pct_match else amount / 100 if abs(amount) <= 100 else 0
    elif "role" in text:
        action_type = "role_change"
    elif "location" in text or "move" in text or "relocate" in text:
        action_type = "location_change"
    elif money_match and any(word in text for word in ["cost", "driver", "benefit", "tax"]):
        action_type = "cost_driver_change"
        cost_delta = float(money_match.group(1).replace(",", ""))

    role = extract_known_value(user_instruction, ["Commercial", "Operations", "Manufacturing", "Finance", "Technology"])
    location = extract_known_value(user_instruction, ["Prague", "Dublin", "Berlin", "Tel Aviv", "London", "New York"])

    return {
        "actions": [
            {
                "type": action_type,
                "month": month,
                "fte_delta": fte_delta,
                "cost_delta": cost_delta,
                "salary_pct_change": salary_pct_change,
                "role": role,
                "location": location,
                "description": user_instruction[:100],
            }
        ]
    }


def extract_known_value(text: str, candidates: List[str]) -> str:
    lower_text = text.lower()
    for candidate in candidates:
        if candidate.lower() in lower_text:
            return candidate
    return ""


def normalize_simulation_response(response: Dict[str, Any], source_text: str) -> Dict[str, Any]:
    actions = response.get("actions", [])
    if not isinstance(actions, list):
        actions = []

    normalized_actions = []
    for action in actions:
        if not isinstance(action, dict):
            continue
        normalized_actions.append(
            {
                "type": str(action.get("type", "cost_driver_change")),
                "month": str(action.get("month", ""))[:7],
                "fte_delta": safe_float(action.get("fte_delta", 0)),
                "cost_delta": safe_float(action.get("cost_delta", 0)),
                "salary_pct_change": safe_float(action.get("salary_pct_change", 0)),
                "role": str(action.get("role", "")),
                "location": str(action.get("location", "")),
                "description": str(action.get("description", source_text))[:100],
            }
        )

    if not normalized_actions:
        return parse_simulation_instruction_fallback(source_text)
    return {"actions": normalized_actions}


def safe_float(value: Any) -> float:
    try:
        return float(value)
    except Exception:
        return 0.0


def build_gpt_context(
    baseline_df: pd.DataFrame,
    drivers_df: pd.DataFrame,
    rules_df: pd.DataFrame,
) -> Dict[str, Any]:
    return {
        "months": sorted(baseline_df["month"].astype(str).unique().tolist()),
        "roles": sorted(baseline_df["role"].astype(str).unique().tolist()),
        "locations": sorted(baseline_df["location"].astype(str).unique().tolist()),
        "drivers": drivers_df.to_dict("records"),
        "rules": rules_df.to_dict("records"),
    }


def apply_simulation_logic(
    baseline_df: pd.DataFrame,
    simulation_response: Dict[str, Any],
    drivers_df: pd.DataFrame,
) -> pd.DataFrame:
    simulation_df = baseline_df.copy()
    if simulation_df.empty:
        return simulation_df

    available_months = sorted(simulation_df["month"].astype(str).unique().tolist())
    default_month = available_months[0]
    benefits_load_pct = get_driver_value(drivers_df, "benefits_load_pct", 0.24)

    for action in simulation_response.get("actions", []):
        action_type = action.get("type", "")
        start_month = action.get("month") or default_month
        role = action.get("role") or None
        location = action.get("location") or None
        fte_delta = safe_float(action.get("fte_delta", 0))
        cost_delta = safe_float(action.get("cost_delta", 0))
        salary_pct_change = safe_float(action.get("salary_pct_change", 0))

        month_mask = simulation_df["month"].astype(str) >= start_month
        role_mask = simulation_df["role"].astype(str).eq(role) if role else pd.Series(True, index=simulation_df.index)
        location_mask = (
            simulation_df["location"].astype(str).eq(location)
            if location
            else pd.Series(True, index=simulation_df.index)
        )
        target_mask = month_mask & role_mask & location_mask
        if not target_mask.any():
            target_mask = month_mask

        target_indices = simulation_df[target_mask].index

        if action_type in ["hire", "layoff"] and fte_delta != 0:
            monthly_count = simulation_df.loc[target_indices].groupby("month").size()
            for idx in target_indices:
                month = simulation_df.at[idx, "month"]
                spread_delta = fte_delta / max(monthly_count.get(month, 1), 1)
                simulation_df.at[idx, "fte"] = max(0, simulation_df.at[idx, "fte"] + spread_delta)
        elif action_type == "salary_change" and salary_pct_change != 0:
            simulation_df.loc[target_indices, "average_salary"] = (
                simulation_df.loc[target_indices, "average_salary"] * (1 + salary_pct_change)
            )
        elif action_type == "cost_driver_change" and cost_delta != 0:
            monthly_count = simulation_df.loc[target_indices].groupby("month").size()
            for idx in target_indices:
                month = simulation_df.at[idx, "month"]
                spread_delta = cost_delta / max(monthly_count.get(month, 1), 1)
                simulation_df.at[idx, "total_cost_of_labor"] = (
                    simulation_df.at[idx, "total_cost_of_labor"] + spread_delta
                )
        elif action_type == "role_change" and role:
            simulation_df.loc[target_indices, "role"] = role
        elif action_type == "location_change" and location:
            simulation_df.loc[target_indices, "location"] = location

        if action_type != "cost_driver_change":
            simulation_df.loc[target_indices, "total_cost_of_labor"] = (
                simulation_df.loc[target_indices, "fte"]
                * simulation_df.loc[target_indices, "average_salary"]
                / 12
                * (1 + benefits_load_pct)
            )

    simulation_df["fte"] = simulation_df["fte"].round(2)
    simulation_df["average_salary"] = simulation_df["average_salary"].round(2)
    simulation_df["total_cost_of_labor"] = simulation_df["total_cost_of_labor"].round(2)
    return simulation_df


def get_driver_value(drivers_df: pd.DataFrame, driver_name: str, default: float) -> float:
    if drivers_df.empty or "driver" not in drivers_df.columns or "value" not in drivers_df.columns:
        return default
    matches = drivers_df[drivers_df["driver"].astype(str).str.lower() == driver_name.lower()]
    if matches.empty:
        return default
    return safe_float(matches.iloc[0]["value"])


def aggregate_monthly(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["month", "fte", "total_cost_of_labor"])
    return (
        df.groupby("month", as_index=False)
        .agg(fte=("fte", "sum"), total_cost_of_labor=("total_cost_of_labor", "sum"))
        .sort_values("month")
    )


def generate_impact_summary(
    baseline_df: pd.DataFrame,
    simulation_df: pd.DataFrame,
    applied_actions: List[Dict[str, Any]],
) -> List[str]:
    baseline_total = aggregate_monthly(baseline_df)
    simulation_total = aggregate_monthly(simulation_df)
    baseline_fte = baseline_total["fte"].iloc[-1] if not baseline_total.empty else 0
    simulation_fte = simulation_total["fte"].iloc[-1] if not simulation_total.empty else 0
    baseline_cost = baseline_total["total_cost_of_labor"].sum() if not baseline_total.empty else 0
    simulation_cost = simulation_total["total_cost_of_labor"].sum() if not simulation_total.empty else 0

    summary = [
        f"FTE impact: {simulation_fte - baseline_fte:+.2f}",
        f"Labor cost impact: ${simulation_cost - baseline_cost:+,.0f}",
    ]
    for action in applied_actions[-4:]:
        summary.append(action.get("description", "Simulation change applied.")[:100])
    return summary[:6]


def render_header() -> None:
    st.markdown(
        f"""
        <div class="app-header">
            <div class="header-logo">
                <img src="{LOGO_URL}" alt="Company logo">
                <span>Teva</span>
            </div>
            <div class="header-title">Workforce Planning Simulation Report</div>
            <div class="header-user">
                <span>TEST USER</span>
                <div class="avatar-placeholder"></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_metric_cards(baseline_df: pd.DataFrame, simulation_df: pd.DataFrame) -> None:
    baseline_monthly = aggregate_monthly(baseline_df)
    simulation_monthly = aggregate_monthly(simulation_df)
    baseline_fte = baseline_monthly["fte"].iloc[-1] if not baseline_monthly.empty else 0
    simulation_fte = simulation_monthly["fte"].iloc[-1] if not simulation_monthly.empty else 0
    baseline_cost = baseline_monthly["total_cost_of_labor"].sum() if not baseline_monthly.empty else 0
    simulation_cost = simulation_monthly["total_cost_of_labor"].sum() if not simulation_monthly.empty else 0

    st.markdown(
        f"""
        <div class="metric-row">
            <div class="metric-card">
                <div class="metric-label">Baseline FTE</div>
                <div class="metric-value">{baseline_fte:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Simulation FTE</div>
                <div class="metric-value">{simulation_fte:,.2f}</div>
            </div>
            <div class="metric-card">
                <div class="metric-label">Total Labor Cost Impact</div>
                <div class="metric-value">${simulation_cost - baseline_cost:+,.0f}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_grouped_column_chart(
    baseline_df: pd.DataFrame,
    simulation_df: pd.DataFrame,
    metric: str,
    title: str,
    y_axis_title: str,
    value_prefix: str = "",
) -> None:
    baseline_monthly = aggregate_monthly(baseline_df)
    simulation_monthly = aggregate_monthly(simulation_df)
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=baseline_monthly["month"],
            y=baseline_monthly[metric],
            name="Baseline",
            marker_color="#2563eb",
            hovertemplate=f"%{{x}}<br>Baseline: {value_prefix}%{{y:,.2f}}<extra></extra>",
        )
    )
    fig.add_trace(
        go.Bar(
            x=simulation_monthly["month"],
            y=simulation_monthly[metric],
            name="Simulation",
            marker_color="#f97316",
            hovertemplate=f"%{{x}}<br>Simulation: {value_prefix}%{{y:,.2f}}<extra></extra>",
        )
    )
    fig.update_layout(
        title={"text": title, "font": {"size": 18, "color": "#172033"}},
        barmode="group",
        height=340,
        margin={"l": 45, "r": 20, "t": 55, "b": 45},
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        legend={"orientation": "h", "yanchor": "bottom", "y": 1.02, "xanchor": "right", "x": 1},
        xaxis={"title": "", "showgrid": False, "tickangle": -30},
        yaxis={"title": y_axis_title, "gridcolor": "#edf0f5", "zerolinecolor": "#d9dee8"},
        font={"family": "Arial, sans-serif", "color": "#27364f"},
    )
    st.plotly_chart(fig, use_container_width=True)


def render_impact_summary(summary: List[str]) -> None:
    st.markdown('<div class="summary-box">', unsafe_allow_html=True)
    st.markdown('<div class="summary-title">Applied Simulation Impact Summary</div>', unsafe_allow_html=True)
    for item in summary[:6]:
        st.markdown(f"- {item[:100]}")
    st.markdown("</div>", unsafe_allow_html=True)


def render_api_connector_placeholder() -> None:
    st.markdown(
        """
        <div class="api-placeholder">
            <strong>Future API connector placeholder</strong><br>
            Ready for HRIS, payroll, finance planning, and data warehouse integrations.
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_chat_panel() -> None:
    st.markdown('<div class="left-panel">', unsafe_allow_html=True)
    st.markdown('<div class="panel-title">Simulation Chat</div>', unsafe_allow_html=True)
    st.markdown(
        """
        <div class="panel-subtitle">
            Use plain language to model hiring, layoffs, salary, role, location,
            and cost-driver changes.
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader("Upload workforce Excel workbook", type=["xlsx"])
    if uploaded_file is not None:
        is_valid, messages = validate_excel_sheets(uploaded_file)
        st.session_state.validation_messages = messages
        if is_valid:
            baseline_df, drivers_df, rules_df = load_excel_data(uploaded_file)
            st.session_state.baseline_df = baseline_df
            st.session_state.drivers_df = drivers_df
            st.session_state.rules_df = rules_df
            st.session_state.simulation_df = baseline_df.copy()
            st.success("Excel data loaded successfully.")
        else:
            for message in messages:
                st.error(message)

    expected_excel_structure()
    st.divider()

    for message in st.session_state.chat_history:
        with st.chat_message(message["role"]):
            st.write(message["content"])

    prompt = st.chat_input("Enter simulation instruction")
    if prompt:
        st.session_state.chat_history.append({"role": "user", "content": prompt})
        context = build_gpt_context(
            st.session_state.baseline_df,
            st.session_state.drivers_df,
            st.session_state.rules_df,
        )
        simulation_response = call_openai_api(prompt, context)
        actions = simulation_response.get("actions", [])
        st.session_state.applied_actions.extend(actions)
        st.session_state.simulation_df = apply_simulation_logic(
            st.session_state.simulation_df,
            simulation_response,
            st.session_state.drivers_df,
        )
        st.session_state.summary = generate_impact_summary(
            st.session_state.baseline_df,
            st.session_state.simulation_df,
            st.session_state.applied_actions,
        )
        assistant_message = "Simulation applied and dashboard updated."
        if actions:
            assistant_message = actions[0].get("description", assistant_message)
        st.session_state.chat_history.append({"role": "assistant", "content": assistant_message})
        st.rerun()

    if st.button("Reset simulation", use_container_width=True):
        st.session_state.simulation_df = st.session_state.baseline_df.copy()
        st.session_state.applied_actions = []
        st.session_state.summary = [
            "No simulation changes applied yet.",
            "FTE impact: 0.00",
            "Labor cost impact: $0",
        ]
        st.session_state.chat_history = [
            {"role": "assistant", "content": "Simulation reset. Enter a new workforce planning change."}
        ]
        st.rerun()

    render_api_connector_placeholder()
    st.markdown("</div>", unsafe_allow_html=True)


def render_report_panel() -> None:
    render_metric_cards(st.session_state.baseline_df, st.session_state.simulation_df)
    render_grouped_column_chart(
        st.session_state.baseline_df,
        st.session_state.simulation_df,
        metric="fte",
        title="FTE Over Time",
        y_axis_title="FTE",
    )
    render_grouped_column_chart(
        st.session_state.baseline_df,
        st.session_state.simulation_df,
        metric="total_cost_of_labor",
        title="Total Cost of Labor Over Time",
        y_axis_title="Total Cost of Labor",
        value_prefix="$",
    )
    render_impact_summary(st.session_state.summary)


def main() -> None:
    inject_css()
    initialize_session_state()
    render_header()
    left_col, right_col = st.columns([1, 4], gap="medium")
    with left_col:
        render_chat_panel()
    with right_col:
        render_report_panel()


if __name__ == "__main__":
    main()
