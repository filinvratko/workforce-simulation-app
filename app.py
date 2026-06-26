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
import streamlit.components.v1 as components


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
    cost_centers = ["CC_001", "CC_002", "CC_003"]
    rows = []

    for cc in cost_centers:
        for employee_id in range(1, 4):
            employee = f"{cc}_E{employee_id:03d}"

            for account in accounts:
                row = {
                    "Company Code": "IL01",
                    "Cost Center": cc,
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
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}"
        )


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
        if str(col).strip().lower() not in excluded_lower
        and is_yyyymm_column(col)
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


def source_to_long_baseline(source_df: pd.DataFrame) -> pd.DataFrame:
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
        var_name="Month",
        value_name="BaselineValue",
    )

    long_df["Month"] = long_df["Month"].astype(str)
    long_df["MonthDate"] = pd.to_datetime(long_df["Month"], format="%Y%m")
    long_df["BaselineValue"] = pd.to_numeric(
        long_df["BaselineValue"], errors="coerce"
    ).fillna(0)

    long_df["SimulationValue"] = long_df["BaselineValue"]
    long_df["IsFTE"] = identify_fte_rows(long_df)

    return long_df


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


def extract_scope(text: str) -> Dict[str, Optional[str]]:
    patterns = [
        ("Cost Center", r"cost\s*center\s+([A-Za-z0-9_\-]+)"),
        ("Company Code", r"company\s*code\s+([A-Za-z0-9_\-]+)"),
        ("Profit Center", r"profit\s*center\s+([A-Za-z0-9_\-]+)"),
        ("Business Area", r"business\s*area\s+([A-Za-z0-9_\-]+)"),
        ("Segment", r"segment\s+([A-Za-z0-9_\-]+)"),
        ("Employee", r"employee\s+([A-Za-z0-9_\-]+)"),
        ("Account", r"account\s+([A-Za-z0-9_\-]+)"),
    ]

    for dimension, pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return {
                "scope_dimension": dimension,
                "scope_value": match.group(1),
            }

    return {
        "scope_dimension": None,
        "scope_value": None,
    }


def fallback_parse_instruction(text: str) -> Dict[str, Any]:
    lower = text.lower()
    month = extract_month_yyyymm(lower)
    scope = extract_scope(text)

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
        "scope_dimension": scope["scope_dimension"],
        "scope_value": scope["scope_value"],
    }

    if any(
        w in lower
        for w in ["hire", "add fte", "increase fte", "recruit", "staff increase"]
    ):
        action["action_type"] = "hire"
        action["fte_delta"] = abs(number)

    elif any(
        w in lower
        for w in ["layoff", "remove fte", "reduce fte", "cut fte", "workforce reduction"]
    ):
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

Rules:
- Baseline is read-only.
- Simulation starts as copy of Baseline.
- Effective date is mandatory.
- Recognize dates like YYYYMM, January 2027, Jan 2027, 2027 January.
- Recognize scoped dimensions such as Cost Center CC_001.

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
      "end_month": null,
      "scope_dimension": "Cost Center" | "Company Code" | "Profit Center" | "Business Area" | "Segment" | "Employee" | "Account" | null,
      "scope_value": string or null
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

        parsed = json.loads(response.choices[0].message.content.strip())

        for action in parsed.get("actions", []):
            if not action.get("scope_dimension"):
                fallback_scope = extract_scope(prompt)
                action["scope_dimension"] = fallback_scope["scope_dimension"]
                action["scope_value"] = fallback_scope["scope_value"]

        return parsed

    except Exception:
        return fallback_parse_instruction(prompt)


def build_scope_mask(detail_df: pd.DataFrame, action: Dict[str, Any]) -> pd.Series:
    mask = pd.Series(True, index=detail_df.index)

    scope_dimension = action.get("scope_dimension")
    scope_value = action.get("scope_value")

    if scope_dimension and scope_value and scope_dimension in detail_df.columns:
        mask &= (
            detail_df[scope_dimension]
            .astype(str)
            .str.strip()
            .str.lower()
            == str(scope_value).strip().lower()
        )

    return mask


def apply_detail_simulation_logic(
    source_df: pd.DataFrame,
    events: List[Dict[str, Any]],
) -> pd.DataFrame:
    detail = source_to_long_baseline(source_df)

    for event in events:
        if event.get("status") != "ready_to_apply":
            continue

        for action in event.get("actions", []):
            effective_month = str(action.get("effective_month", ""))

            if not re.fullmatch(r"\d{6}", effective_month):
                continue

            date_mask = detail["Month"] >= effective_month
            scope_mask = build_scope_mask(detail, action)
            target_mask = date_mask & scope_mask

            action_type = action.get("action_type")
            fte_delta = float(action.get("fte_delta", 0) or 0)
            cost_pct_delta = float(action.get("cost_pct_delta", 0) or 0)
            cost_abs_delta = float(action.get("cost_abs_delta", 0) or 0)

            if action_type in ["hire", "layoff"]:
                fte_mask = target_mask & detail["IsFTE"]
                rows_per_month = detail.loc[fte_mask].groupby("Month").size().to_dict()

                for month, row_count in rows_per_month.items():
                    month_mask = fte_mask & (detail["Month"] == month)
                    detail.loc[month_mask, "SimulationValue"] += fte_delta / row_count

                cost_mask = target_mask & ~detail["IsFTE"]
                affected_cost = detail.loc[cost_mask, "BaselineValue"].sum()
                affected_fte = detail.loc[fte_mask, "BaselineValue"].sum()

                if affected_fte != 0:
                    cost_per_fte = affected_cost / affected_fte
                    cost_delta_total = fte_delta * cost_per_fte
                    cost_rows_per_month = (
                        detail.loc[cost_mask].groupby("Month").size().to_dict()
                    )

                    for month, row_count in cost_rows_per_month.items():
                        month_mask = cost_mask & (detail["Month"] == month)
                        detail.loc[month_mask, "SimulationValue"] += (
                            cost_delta_total / row_count
                        )

            elif action_type == "cost_change":
                cost_mask = target_mask & ~detail["IsFTE"]
                detail.loc[cost_mask, "SimulationValue"] *= 1 + cost_pct_delta
                detail.loc[cost_mask, "SimulationValue"] += cost_abs_delta

    detail["SimulationValue"] = detail["SimulationValue"].clip(lower=0)
    detail["Delta"] = detail["SimulationValue"] - detail["BaselineValue"]

    return detail


def aggregate_monthly_from_detail(detail_df: pd.DataFrame) -> pd.DataFrame:
    fte_df = (
        detail_df[detail_df["IsFTE"]]
        .groupby("Month", as_index=False)
        .agg(
            BaselineFTE=("BaselineValue", "sum"),
            SimulationFTE=("SimulationValue", "sum"),
        )
    )

    cost_df = (
        detail_df[~detail_df["IsFTE"]]
        .groupby("Month", as_index=False)
        .agg(
            BaselineUSD=("BaselineValue", "sum"),
            SimulationUSD=("SimulationValue", "sum"),
        )
    )

    monthly = pd.merge(fte_df, cost_df, on="Month", how="outer").fillna(0)
    monthly["MonthDate"] = pd.to_datetime(monthly["Month"], format="%Y%m")
    monthly = monthly.sort_values("MonthDate")
    monthly["MonthLabel"] = monthly["MonthDate"].dt.strftime("%b %Y")

    return monthly


def build_simulation_detail(detail_df: pd.DataFrame) -> pd.DataFrame:
    ordered_cols = DATA_SAMPLE_ROW_COLUMNS + [
        "Month",
        "BaselineValue",
        "SimulationValue",
        "Delta",
    ]

    result = detail_df[ordered_cols].copy()
    result["Month"] = result["Month"].astype(str)

    return result.sort_values(DATA_SAMPLE_ROW_COLUMNS + ["Month"])


def reset_simulation_only() -> None:
    st.session_state.simulation_events = []
    st.session_state.simulation_requests = []
    st.session_state.simulation_input_counter += 1
    st.session_state.force_reset_simulation = True


def render_chart(title, df, baseline_metric, simulation_metric, y_title, height):
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
            x=1,
        ),
        yaxis_title=y_title,
    )

    fig.update_yaxes(gridcolor="#374151")
    fig.update_xaxes(showgrid=False, tickangle=-45)

    return fig


def render_datapoints_row(
    baseline_avg_fte: float,
    simulation_avg_fte: float,
    baseline_end_fte: float,
    simulation_end_fte: float,
    baseline_total_usd: float,
    simulation_total_usd: float,
) -> None:
    fte_delta = simulation_avg_fte - baseline_avg_fte
    cost_delta = simulation_total_usd - baseline_total_usd

    fte_delta_class = "positive" if fte_delta >= 0 else "negative"
    cost_delta_class = "positive" if cost_delta >= 0 else "negative"

    fte_arrow = "↑" if fte_delta >= 0 else "↓"
    cost_arrow = "↑" if cost_delta >= 0 else "↓"

    html_block = f"""
    <style>
        .dp-grid {{
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 18px;
            width: 100%;
        }}
        .dp-card {{
            background: #111827;
            border: 1px solid #1F2937;
            border-radius: 10px;
            height: 126px;
            padding: 18px 22px;
            box-sizing: border-box;
            overflow: hidden;
            font-family: sans-serif;
        }}
        .dp-title {{
            color: #F9FAFB;
            font-size: 14px;
            font-weight: 700;
            margin-bottom: 14px;
            white-space: nowrap;
        }}
        .dp-main-area {{
            display: flex;
            align-items: center;
            justify-content: flex-start;
            gap: 18px;
            width: 100%;
        }}
        .dp-main-stack {{
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: flex-start;
            width: fit-content;
            max-width: 78%;
        }}
        .dp-main {{
            color: #F9FAFB;
            font-size: 34px;
            line-height: 38px;
            font-weight: 800;
            white-space: nowrap;
            text-align: center;
        }}
        .dp-secondary {{
            color: #CBD5E1;
            font-size: 10px;
            line-height: 12px;
            margin-top: 8px;
            text-align: center;
            white-space: nowrap;
        }}
        .dp-delta {{
            align-self: center;
            font-size: 14px;
            line-height: 18px;
            font-weight: 800;
            white-space: nowrap;
            padding: 5px 10px;
            border-radius: 999px;
        }}
        .dp-delta.positive {{
            color: #22C55E;
            background: rgba(34, 197, 94, 0.15);
        }}
        .dp-delta.negative {{
            color: #EF4444;
            background: rgba(239, 68, 68, 0.15);
        }}
    </style>

    <div class="dp-grid">
        <div class="dp-card">
            <div class="dp-title">Baseline FTE</div>
            <div class="dp-main-area">
                <div class="dp-main-stack">
                    <div class="dp-main">{baseline_avg_fte:,.1f}</div>
                    <div class="dp-secondary">FTE at end of period: <b>{baseline_end_fte:,.1f}</b></div>
                </div>
            </div>
        </div>

        <div class="dp-card">
            <div class="dp-title">Simulation FTE</div>
            <div class="dp-main-area">
                <div class="dp-main-stack">
                    <div class="dp-main">{simulation_avg_fte:,.1f}</div>
                    <div class="dp-secondary">FTE at end of period: <b>{simulation_end_fte:,.1f}</b></div>
                </div>
                <div class="dp-delta {fte_delta_class}">{fte_arrow} {fte_delta:+,.1f}</div>
            </div>
        </div>

        <div class="dp-card">
            <div class="dp-title">Baseline Labor Cost</div>
            <div class="dp-main-area">
                <div class="dp-main-stack">
                    <div class="dp-main">${baseline_total_usd:,.0f}</div>
                </div>
            </div>
        </div>

        <div class="dp-card">
            <div class="dp-title">Simulation Labor Cost</div>
            <div class="dp-main-area">
                <div class="dp-main-stack">
                    <div class="dp-main">${simulation_total_usd:,.0f}</div>
                </div>
                <div class="dp-delta {cost_delta_class}">{cost_arrow} ${cost_delta:+,.0f}</div>
            </div>
        </div>
    </div>
    """

    components.html(html_block, height=132, scrolling=False)


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
        "force_reset_simulation": False,
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
                st.session_state.force_reset_simulation = True
                st.session_state.last_data_sample_upload_id = upload_id
                st.rerun()

        if st.button("Reset baseline to mock data", use_container_width=True):
            st.session_state.data_sample_df = generate_mock_source_data()
            st.session_state.data_sample_source = "Mock data"
            st.session_state.simulation_events = []
            st.session_state.simulation_requests = []
            st.session_state.force_reset_simulation = True
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

        st.divider()

        input_key = f"simulation_input_text_{st.session_state.simulation_input_counter}"

        st.markdown('<div class="box">', unsafe_allow_html=True)
        simulation_prompt = st.text_area(
            "Simulation input",
            placeholder="Example: Increase FTE by 10 for Cost Center CC_001 from January 2027",
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
                st.session_state.force_reset_simulation = False

            st.session_state.simulation_requests.insert(
                0,
                {
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "request": prompt,
                    "response": parsed.get("executive_summary", ""),
                    "status": parsed.get("status", ""),
                },
            )

            st.session_state.simulation_input_counter += 1
            st.rerun()

        st.button(
            "Reset simulation only",
            use_container_width=True,
            on_click=reset_simulation_only,
        )

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

    try:
        active_events = (
            []
            if st.session_state.force_reset_simulation
            else st.session_state.simulation_events
        )

        detail_df = apply_detail_simulation_logic(
            st.session_state.data_sample_df,
            active_events,
        )

        simulation_df = aggregate_monthly_from_detail(detail_df)
        simulation_detail_df = build_simulation_detail(detail_df)

        if st.session_state.force_reset_simulation:
            st.session_state.force_reset_simulation = False

    except Exception as exc:
        with right:
            st.error(f"Data preparation error: {exc}")
        return

    baseline_avg_fte = simulation_df["BaselineFTE"].mean()
    simulation_avg_fte = simulation_df["SimulationFTE"].mean()
    baseline_end_fte = simulation_df["BaselineFTE"].iloc[-1]
    simulation_end_fte = simulation_df["SimulationFTE"].iloc[-1]

    baseline_total_usd = simulation_df["BaselineUSD"].sum()
    simulation_total_usd = simulation_df["SimulationUSD"].sum()

    with right:
        render_datapoints_row(
            baseline_avg_fte,
            simulation_avg_fte,
            baseline_end_fte,
            simulation_end_fte,
            baseline_total_usd,
            simulation_total_usd,
        )

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

        for item in build_summary(simulation_df, active_events):
            st.markdown(f"- {item}")

        st.markdown("</div>", unsafe_allow_html=True)

        with st.expander("Simulation Detail"):
            st.dataframe(simulation_detail_df, use_container_width=True)

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
