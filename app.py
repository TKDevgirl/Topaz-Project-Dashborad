import streamlit as st

from auth import render_sidebar
from config import APP_ICON, APP_TITLE
from dashboard import render_dashboard, render_hero
from excel_reader import empty_report_df, generate_report
from style import apply_style

st.set_page_config(
    page_title=APP_TITLE,
    page_icon=APP_ICON,
    layout="wide",
)

apply_style()
render_sidebar()

defaults = {
    "result_df": empty_report_df(),
    "has_result": False,
    "report": None,
    "total_docs": 0,
    "focus_docs": 0,
    "dashboard_matrix": None,
    "action_counts": {},
    "last_updated": None,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

render_hero()

if st.session_state.role == "admin":
    st.markdown(
        '<div class="admin">👩‍💼 Admin mode: Admin can upload files and generate dashboard.</div>',
        unsafe_allow_html=True,
    )

    upload_col_1, upload_col_2, button_col = st.columns([1, 1, 0.8])

    with upload_col_1:
        tracking_file = st.file_uploader("1) Upload Tracking_document.xlsx", type=["xlsx"])

    with upload_col_2:
        takenaka_file = st.file_uploader("2) Upload Takenaka Summary.xlsx", type=["xlsx"])

    with button_col:
        st.write("")
        st.write("")
        generate_btn = st.button("🚀 Generate Dashboard", type="primary", use_container_width=True)

    if tracking_file and takenaka_file and generate_btn:
        with st.spinner("Reading files and generating dashboard..."):
            result = generate_report(tracking_file, takenaka_file)

        st.session_state.result_df = result["df"]
        st.session_state.has_result = True
        st.session_state.report = result["report"]
        st.session_state.total_docs = result["total_docs"]
        st.session_state.focus_docs = result["focus_docs"]
        st.session_state.dashboard_matrix = result["dashboard_matrix"]
        st.session_state.action_counts = result["action_counts"]
        st.session_state.last_updated = result["last_updated"]

        st.rerun()

else:
    st.markdown(
        '<div class="notice">ℹ️ Viewer mode: only Admin can upload files and generate dashboard.</div>',
        unsafe_allow_html=True,
    )

render_dashboard()
