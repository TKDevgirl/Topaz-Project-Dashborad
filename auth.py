import streamlit as st
from config import ADMIN_USERS, APP_VERSION


def init_auth_state():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if "role" not in st.session_state:
        st.session_state.role = "viewer"

    if "username" not in st.session_state:
        st.session_state.username = ""


def render_sidebar():
    init_auth_state()

    with st.sidebar:
        st.markdown('<div class="logo">💎 TOPAZ</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="logo-sub">{APP_VERSION}</div>', unsafe_allow_html=True)

        if not st.session_state.logged_in:
            username = st.text_input("Username")

            if st.button("Login", use_container_width=True):
                clean_username = username.strip() or "viewer"
                st.session_state.username = clean_username
                st.session_state.role = "admin" if clean_username.lower() in ADMIN_USERS else "viewer"
                st.session_state.logged_in = True
                st.rerun()

        else:
            st.markdown(
                f"""
                <div class="user-card">
                    <b>👤 User</b><br>{st.session_state.username}<br><br>
                    <b>🔑 Role</b><br>{st.session_state.role.title()}
                </div>
                """,
                unsafe_allow_html=True,
            )

            if st.button("Logout", use_container_width=True):
                st.session_state.logged_in = False
                st.session_state.role = "viewer"
                st.session_state.username = ""
                st.rerun()

        st.divider()
        st.markdown('<div class="nav-active">🏠 Dashboard</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📋 Documents</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">📊 Action Summary</div>', unsafe_allow_html=True)
        st.markdown('<div class="nav-item">⬇ Export Report</div>', unsafe_allow_html=True)
