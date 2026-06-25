def render_custom_metric(
    title: str,
    main_value: str,
    sub_label: str = "",
    sub_value: str = "",
    delta: str = "",
    delta_positive: bool = True,
) -> None:
    delta_class = "positive" if delta_positive else "negative"

    sub_html = ""
    if sub_label:
        sub_html = f"""
        <div class="metric-sub">
            {sub_label}: <b>{sub_value}</b>
        </div>
        """

    delta_html = ""
    if delta:
        delta_html = f"""
        <div class="metric-delta {delta_class}">
            {delta}
        </div>
        """

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
