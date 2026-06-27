import streamlit as st


def apply_style():
    st.markdown("""
    <style>
    .stApp {
        background: #f4f7fb;
        color: #0f172a;
    }

    .block-container {
        max-width: 1500px;
        padding-top: 1.2rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #061124 0%, #0f172a 72%, #111827 100%);
    }

    [data-testid="stSidebar"] * {
        color: white;
    }

    .logo {
        font-size: 30px;
        font-weight: 900;
    }

    .logo-sub {
        color: #c7d2fe;
        font-size: 13px;
        margin-bottom: 20px;
    }

    .user-card {
        padding: 16px;
        border-radius: 18px;
        background: rgba(255,255,255,.08);
        border: 1px solid rgba(255,255,255,.14);
        margin-bottom: 18px;
    }

    .nav-item {
        padding: 10px 12px;
        border-radius: 12px;
        margin: 6px 0;
        background: rgba(255,255,255,.05);
        font-weight: 700;
    }

    .nav-active {
        padding: 10px 12px;
        border-radius: 12px;
        margin: 6px 0;
        background: linear-gradient(90deg,#4f46e5,#7c3aed);
        font-weight: 800;
    }

    .hero {
        background: linear-gradient(90deg,#071226 0%, #0b1b4d 50%, #172554 100%);
        color: white;
        border-radius: 24px;
        padding: 28px 32px;
        box-shadow: 0 20px 40px rgba(15,23,42,.20);
        margin-bottom: 18px;
    }

    .hero-title {
        font-size: 38px;
        font-weight: 900;
        margin-bottom: 6px;
    }

    .hero-sub {
        color:#dbeafe;
        font-size:15px;
    }

    .admin {
        padding: 15px 18px;
        border-radius: 18px;
        margin: 16px 0 18px 0;
        font-weight: 700;
        background: linear-gradient(90deg,#ecfdf5,#ffffff);
        border:1px solid #bbf7d0;
        color:#166534;
    }

    .notice {
        padding: 15px 18px;
        border-radius: 18px;
        margin: 16px 0 18px 0;
        font-weight: 700;
        background: linear-gradient(90deg,#eff6ff,#ffffff);
        border:1px solid #bfdbfe;
        color:#1d4ed8;
    }

    .kpi-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 22px;
        padding: 22px;
        box-shadow: 0 14px 32px rgba(15,23,42,.08);
        min-height: 145px;
        position: relative;
        overflow:hidden;
    }

    .kpi-card:after {
        content:"";
        position:absolute;
        left:0;
        right:0;
        bottom:0;
        height:5px;
        background:var(--accent);
    }

    .kpi-title {
        color:#64748b;
        font-size:13px;
        font-weight:850;
    }

    .kpi-value {
        font-size: 38px;
        font-weight: 950;
        color:#0f172a;
        line-height:1.05;
        margin:7px 0;
    }

    .kpi-sub {
        color:#64748b;
        font-size:13px;
    }

    .panel {
        background: white;
        border:1px solid #e2e8f0;
        border-radius: 22px;
        box-shadow: 0 14px 32px rgba(15,23,42,.07);
        padding: 22px;
        margin-bottom: 18px;
    }

    .panel-title {
        font-size: 20px;
        font-weight: 900;
        color:#0f172a;
        margin-bottom: 14px;
    }

    .quick-row {
        display:flex;
        justify-content:space-between;
        align-items:center;
        padding:12px;
        border-radius:14px;
        background:#f8fafc;
        border:1px solid #e2e8f0;
        margin-bottom:8px;
    }

    .count-pill {
        padding:4px 10px;
        border-radius:999px;
        font-weight:900;
        background:#dbeafe;
        color:#1d4ed8;
    }
    </style>
    """, unsafe_allow_html=True)
