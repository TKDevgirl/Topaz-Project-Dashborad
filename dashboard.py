import altair as alt
import pandas as pd
import streamlit as st

from excel_reader import empty_report_df


def render_hero():
    st.markdown(
        """
        <div class="hero">
            <div class="hero-grid">
                <div style="font-size:42px;">💎</div>
                <div>
                    <div class="hero-title">Topaz Smart Document Tracker</div>
                    <div class="hero-sub">TOPAZ BKK1 | ICT Document Control Dashboard</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_summary_matrix(matrix):
    st.markdown('<div class="panel"><div class="panel-title">📌 Dashboard Summary</div>', unsafe_allow_html=True)
    if matrix is not None and not matrix.empty:
        st.dataframe(matrix.set_index("Status"), use_container_width=True)
    else:
        st.info("Dashboard sheet summary not found.")
    st.markdown("</div>", unsafe_allow_html=True)


def render_kpi_cards(total_docs, focus_docs, action_counts):
    cards = [
        ("Total Documents", total_docs, "From Dashboard sheet", "#7c3aed"),
        ("Open / On Progress", focus_docs, "Open + on progress", "#2563eb"),
        ("Open & On Process", action_counts.get("OPEN & ON PROCESS", 0), "Compared with Takenaka", "#16a34a"),
        ("Need Update", action_counts.get("UPDATE TRACKING TO CLOSED", 0), "Update tracking", "#f97316"),
        ("Overdue", action_counts.get("OVERDUE / FOLLOW UP", 0), "Follow up", "#ef4444"),
    ]

    columns = st.columns(5)
    for col, (title, value, sub, accent) in zip(columns, cards):
        with col:
            st.markdown(f"""<div class="kpi-card" style="--accent:{accent};"><div class="kpi-title">{title}</div><div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)


def render_action_summary(action_counts):
    st.markdown('<div class="panel"><div class="panel-title">📊 Action Summary</div>', unsafe_allow_html=True)
    summary_df = pd.DataFrame([{"Action": key, "Count": value} for key, value in action_counts.items()])
    if not summary_df.empty:
        chart = alt.Chart(summary_df).mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8).encode(
            x=alt.X("Action:N", sort="-y", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("Count:Q"),
            tooltip=["Action", "Count"],
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)
        st.dataframe(summary_df, use_container_width=True, hide_index=True)
    else:
        st.info("No action data found.")
    st.markdown("</div>", unsafe_allow_html=True)


def render_quick_action(action_counts):
    st.markdown('<div class="panel"><div class="panel-title">⚡ Quick Action</div>', unsafe_allow_html=True)
    quick_items = [
        ("Returned by NV5", action_counts.get("RETURNED BY NV5 / NEED RESUBMIT", 0)),
        ("Need update closed", action_counts.get("UPDATE TRACKING TO CLOSED", 0)),
        ("Overdue follow up", action_counts.get("OVERDUE / FOLLOW UP", 0)),
        ("Not found in Takenaka", action_counts.get("NOT FOUND IN TAKENAKA SOURCE", 0)),
        ("Check manually", action_counts.get("CHECK", 0)),
        ("Open only", action_counts.get("OPEN", 0)),
    ]
    for label, count in quick_items:
        st.markdown(f'<div class="quick-row"><span>{label}</span><span class="count-pill">{count}</span></div>', unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_document_table(df, report):
    if df is None or df.empty:
        df = empty_report_df()
    if "Action" not in df.columns:
        df["Action"] = ""

    st.markdown('<div class="panel"><div class="panel-title">📋 Document Action List</div>', unsafe_allow_html=True)

    filter_col, search_col, export_col = st.columns([1, 2, 0.8])
    actions = sorted([item for item in df["Action"].fillna("").astype(str).unique().tolist() if item])

    with filter_col:
        selected_action = st.selectbox("Filter by Action", ["All"] + actions)
    with search_col:
        search = st.text_input("Search Document No / Document Name")
    with export_col:
        st.write("")
        st.write("")
        if report is not None:
            st.download_button(
                label="⬇️ Export Excel",
                data=report,
                file_name="Open_On_Process_Compare.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    filtered_df = df.copy()
    if selected_action != "All":
        filtered_df = filtered_df[filtered_df["Action"] == selected_action]
    if search:
        filtered_df = filtered_df[
            filtered_df.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
        ]

    st.dataframe(filtered_df, use_container_width=True, height=480)
    st.markdown("</div>", unsafe_allow_html=True)


def render_dashboard():
    if not st.session_state.get("has_result", False):
        st.info("Please upload both Excel files and click Generate Dashboard.")
        return

    df = st.session_state.get("result_df", empty_report_df())
    matrix = st.session_state.get("dashboard_matrix")
    action_counts = st.session_state.get("action_counts", {})
    report = st.session_state.get("report")

    st.success(f"Dashboard generated successfully ✅ | Last updated: {st.session_state.get('last_updated')}")

    render_kpi_cards(
        total_docs=st.session_state.get("total_docs", 0),
        focus_docs=st.session_state.get("focus_docs", 0),
        action_counts=action_counts,
    )
    st.write("")

    render_summary_matrix(matrix)

    chart_col, quick_col = st.columns([2, 1])
    with chart_col:
        render_action_summary(action_counts)
    with quick_col:
        render_quick_action(action_counts)

    render_document_table(df, report)
