import streamlit as st

APP_VERSION = "1.0.3"

# ---------------------------------------------------------------------------
# PAGE CONFIG & THEME
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title='NNS Evidence Report Generator',
    page_icon='📋',
    layout='wide',
    initial_sidebar_state='collapsed',
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@400;600;700;800&display=swap');

:root {
    --bg: #0a0a0a;
    --surface: #111111;
    --surface2: #1a1a1a;
    --border: #2a2a2a;
    --accent: #f97316;
    --accent2: #ea580c;
    --success: #22c55e;
    --warn: #f59e0b;
    --danger: #ef4444;
    --text: #f0ece8;
    --text-muted: #6b6b6b;
    --mono: 'DM Mono', monospace;
    --sans: 'Syne', sans-serif;
}

html, body, [data-testid="stAppViewContainer"] {
    background-color: var(--bg) !important;
    color: var(--text);
    font-family: var(--sans);
}

[data-testid="stAppViewContainer"] > .main {
    background-color: var(--bg);
}

[data-testid="stHeader"] { background: transparent !important; }
[data-testid="stToolbar"] { display: none; }
[data-testid="stSidebar"] { background: var(--surface) !important; }
[data-testid="stDecoration"] { display: none; }

/* Hide Streamlit branding */
#MainMenu { visibility: hidden; }

/* App footer */
.app-footer {
    position: fixed;
    bottom: 0; left: 0; right: 0;
    background: var(--surface);
    border-top: 1px solid var(--border);
    padding: 0 2rem;
    height: 2.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
    z-index: 999;
    gap: 1.5rem;
    line-height: 1;
}
.footer-logo {
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 800;
    letter-spacing: 0.12em;
    color: var(--text);
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 4px;
    padding: 0.2rem 0.55rem;
    display: flex;
    align-items: center;
    line-height: 1;
    flex-shrink: 0;
    height: 1.5rem;
    box-sizing: border-box;
}
.footer-logo span { color: var(--accent); }
.footer-disclaimer {
    font-family: var(--mono);
    font-size: 0.62rem;
    color: var(--text-muted);
    text-align: center;
    line-height: 1;
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
}
.footer-version {
    font-family: var(--mono);
    font-size: 0.62rem;
    color: var(--text-muted);
    white-space: nowrap;
    display: flex;
    align-items: center;
    line-height: 1;
    flex-shrink: 0;
}
/* Offset main content so footer doesn't overlap it */
[data-testid="stAppViewContainer"] > .main > .block-container {
    padding-bottom: 4rem !important;
}

/* Typography */
h1, h2, h3, h4 { font-family: var(--sans); font-weight: 800; color: var(--text); }

/* Hero */
.hero {
    background: linear-gradient(135deg, #0a0a0a 0%, #111111 50%, #1a0f00 100%);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -60px; right: -60px;
    width: 200px; height: 200px;
    background: radial-gradient(circle, rgba(249,115,22,0.14) 0%, transparent 70%);
    border-radius: 50%;
}
.hero::after {
    content: '';
    position: absolute;
    bottom: -40px; left: 30%;
    width: 150px; height: 150px;
    background: radial-gradient(circle, rgba(234,88,12,0.08) 0%, transparent 70%);
    border-radius: 50%;
}
.hero-tag {
    font-family: var(--mono);
    font-size: 0.7rem;
    letter-spacing: 0.15em;
    color: var(--accent);
    text-transform: uppercase;
    margin-bottom: 0.75rem;
}
.hero h1 {
    font-size: 2.2rem;
    font-weight: 800;
    margin: 0 0 0.5rem 0;
    line-height: 1.15;
    background: linear-gradient(135deg, #f0ece8 0%, #d4a87a 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.hero p {
    color: var(--text-muted);
    font-size: 0.95rem;
    margin: 0;
    font-family: var(--mono);
}

/* Section cards */
.section-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.5rem;
    margin-bottom: 1.25rem;
}
.section-label {
    font-family: var(--mono);
    font-size: 0.65rem;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 0.75rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.section-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
}

/* Region badge */
.region-badge {
    display: inline-block;
    padding: 0.2rem 0.65rem;
    border-radius: 999px;
    font-family: var(--mono);
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: 0.08em;
}
.region-mcc { background: rgba(249,115,22,0.12); color: #f97316; border: 1px solid rgba(249,115,22,0.35); }
.region-cs  { background: rgba(234,88,12,0.12);  color: #fb923c; border: 1px solid rgba(234,88,12,0.35); }

/* Inputs */
[data-testid="stTextInput"] input,
[data-testid="stSelectbox"] > div > div,
[data-testid="stTextArea"] textarea {
    background: var(--surface2) !important;
    border: 1px solid var(--border) !important;
    color: var(--text) !important;
    border-radius: 8px !important;
    font-family: var(--mono) !important;
    font-size: 0.875rem !important;
}
[data-testid="stTextInput"] input:focus,
[data-testid="stTextArea"] textarea:focus {
    border-color: var(--accent) !important;
    box-shadow: 0 0 0 2px rgba(249,115,22,0.20) !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1px dashed var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--accent) !important;
    background: rgba(249,115,22,0.04) !important;
}

/* Buttons */
[data-testid="stButton"] button {
    background: var(--accent) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: var(--sans) !important;
    font-weight: 700 !important;
    font-size: 0.9rem !important;
    padding: 0.6rem 1.5rem !important;
    transition: all 0.2s !important;
}
[data-testid="stButton"] button:hover {
    background: #ea580c !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 20px rgba(249,115,22,0.4) !important;
}

/* Secondary / Clear buttons — red */
[data-testid="stButton"] button[kind="secondary"] {
    background: rgba(239,68,68,0.1) !important;
    color: #ef4444 !important;
    border: 1px solid rgba(239,68,68,0.4) !important;
    font-weight: 600 !important;
}
[data-testid="stButton"] button[kind="secondary"]:hover {
    background: rgba(239,68,68,0.2) !important;
    color: #fca5a5 !important;
    border-color: #ef4444 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 16px rgba(239,68,68,0.25) !important;
}

/* Download button */
[data-testid="stDownloadButton"] button {
    background: linear-gradient(135deg, var(--success), #16a34a) !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    font-family: var(--sans) !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    padding: 0.75rem 2rem !important;
    width: 100% !important;
    transition: all 0.2s !important;
}
[data-testid="stDownloadButton"] button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 20px rgba(34,197,94,0.3) !important;
}

/* Tags row */
.tag-row { display: flex; flex-wrap: wrap; gap: 0.4rem; margin-top: 0.5rem; }
.tag {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 0.2rem 0.6rem;
    font-family: var(--mono);
    font-size: 0.75rem;
    color: var(--text-muted);
}
.tag-accent { border-color: var(--accent); color: var(--accent); background: rgba(249,115,22,0.08); }

/* Divider */
.hr { border: none; border-top: 1px solid var(--border); margin: 1.5rem 0; }

/* Status indicators */
.status-row {
    display: flex; align-items: center; gap: 0.5rem;
    font-family: var(--mono); font-size: 0.8rem; color: var(--text-muted);
    padding: 0.4rem 0;
}
.dot { width: 6px; height: 6px; border-radius: 50%; flex-shrink: 0; }
.dot-green { background: var(--success); }
.dot-yellow { background: var(--warn); }
.dot-red { background: var(--danger); }

/* Summary result card */
.result-card {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 10px;
    padding: 1.25rem 1.5rem;
    margin-bottom: 0.75rem;
}
.result-grid {
    display: grid;
    grid-template-columns: repeat(4, 1fr);
    gap: 1rem;
    margin-top: 1rem;
}
.metric-box {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 8px;
    padding: 0.75rem 1rem;
    text-align: center;
}
.metric-val {
    font-family: var(--mono);
    font-size: 1.5rem;
    font-weight: 500;
    color: var(--accent);
}
.metric-lbl {
    font-family: var(--mono);
    font-size: 0.65rem;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-top: 0.2rem;
}

/* Alerts */
.alert {
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-family: var(--mono);
    font-size: 0.8rem;
    margin: 0.5rem 0;
}
.alert-warn { background: rgba(245,158,11,0.1); border: 1px solid rgba(245,158,11,0.3); color: #fbbf24; }
.alert-success { background: rgba(34,197,94,0.1); border: 1px solid rgba(34,197,94,0.3); color: #4ade80; }
.alert-info { background: rgba(249,115,22,0.1); border: 1px solid rgba(249,115,22,0.3); color: #fdba74; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# CUSTOMER TABS
# ---------------------------------------------------------------------------

from sketchup.app_sketchup import render as render_sketchup
from bentley.app_bentley   import render as render_bentley

tab_sketchup, tab_bentley = st.tabs([
    "🟦 SketchUp · Trimble",
    "🟫 Bentley",
])

with tab_sketchup:
    st.markdown("""
    <div class="hero">
        <div class="hero-tag">▸ Trimble SketchUp · License Compliance</div>
        <h1>Evidence Report Generator</h1>
        <p>Automated evidence report generation for MCC and Cono Sur regions</p>
    </div>
    """, unsafe_allow_html=True)
    render_sketchup()

with tab_bentley:
    st.markdown("""
    <div class="hero">
        <div class="hero-tag">▸ Bentley · License Compliance</div>
        <h1>Evidence Report Generator</h1>
        <p>NNS evidence report generation for Bentley products</p>
    </div>
    """, unsafe_allow_html=True)
    render_bentley()

# APP FOOTER (always visible)
# ---------------------------------------------------------------------------
st.markdown(f"""
<div class="app-footer">
    <div class="footer-logo">RUVI<span>XX</span></div>
    <div class="footer-disclaimer">
        ⚠ Prototype — Internal Use Only &nbsp;|&nbsp;
        Automates the NNS Evidence Report process for LATAM at Ruvixx &nbsp;|&nbsp;
        All data and results are considered <strong>confidential</strong>
    </div>
    <div class="footer-version">v{APP_VERSION}</div>
</div>
""", unsafe_allow_html=True)

