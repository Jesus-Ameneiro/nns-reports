import io
import json
import openpyxl
import pandas as pd
import streamlit as st
from pathlib import Path
from processor import run_processing
from report_writer import fill_template, patch_and_save

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

APP_VERSION = "1.0.3"
CONFIG_PATH = Path(__file__).parent / 'config.json'

@st.cache_data
def load_config():
    with open(CONFIG_PATH) as f:
        return json.load(f)

def get_region(country, config):
    for region, data in config['regions'].items():
        if country in data['countries']:
            return region
    return None

def all_countries(config):
    countries = []
    for region, data in config['regions'].items():
        for c in data['countries']:
            countries.append((c, region, data['name']))
    return sorted(countries, key=lambda x: x[0])

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
    --bg: #0d0f14;
    --surface: #13161e;
    --surface2: #1a1e28;
    --border: #252a38;
    --accent: #4f8ef7;
    --accent2: #7c3aed;
    --success: #22c55e;
    --warn: #f59e0b;
    --danger: #ef4444;
    --text: #e2e8f0;
    --text-muted: #64748b;
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
    background: linear-gradient(135deg, #0d0f14 0%, #13161e 50%, #1a1228 100%);
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
    background: radial-gradient(circle, rgba(79,142,247,0.12) 0%, transparent 70%);
    border-radius: 50%;
}
.hero::after {
    content: '';
    position: absolute;
    bottom: -40px; left: 30%;
    width: 150px; height: 150px;
    background: radial-gradient(circle, rgba(124,58,237,0.08) 0%, transparent 70%);
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
    background: linear-gradient(135deg, #e2e8f0 0%, #94a3b8 100%);
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
.region-mcc { background: rgba(79,142,247,0.15); color: #4f8ef7; border: 1px solid rgba(79,142,247,0.3); }
.region-cs  { background: rgba(124,58,237,0.15); color: #a78bfa; border: 1px solid rgba(124,58,237,0.3); }

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
    box-shadow: 0 0 0 2px rgba(79,142,247,0.15) !important;
}

/* File uploader */
[data-testid="stFileUploader"] {
    background: var(--surface2) !important;
    border: 1px dashed var(--border) !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--accent) !important;
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
    background: #6aa0f8 !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 20px rgba(79,142,247,0.3) !important;
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
.tag-accent { border-color: var(--accent); color: var(--accent); background: rgba(79,142,247,0.08); }

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
.alert-info { background: rgba(79,142,247,0.1); border: 1px solid rgba(79,142,247,0.3); color: #93c5fd; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# HEADER
# ---------------------------------------------------------------------------

st.markdown("""
<div class="hero">
    <div class="hero-tag">▸ Trimble SketchUp · License Compliance</div>
    <h1>Evidence Report Generator</h1>
    <p>Automated evidence report generation for MCC and Cono Sur regions</p>
</div>
""", unsafe_allow_html=True)

config = load_config()
countries_list = all_countries(config)

# ---------------------------------------------------------------------------
# LAYOUT: Two columns — inputs left, files right
# ---------------------------------------------------------------------------

left_col, right_col = st.columns([1, 1], gap='large')

# ===========================================================================
# LEFT COLUMN — Case Information
# ===========================================================================

with left_col:

    # --- CASE INFORMATION ---
    st.markdown('<div class="section-label">01 · Case Information</div>', unsafe_allow_html=True)

    entity_name = st.text_input(
        'Entity / Organization Name',
        placeholder='e.g. Panificadora la Esperanza S.A. DE C.V.',
        key='entity_name'
    )

    # Case IDs — dynamic list
    st.markdown('<div style="font-size:0.85rem; color:#94a3b8; margin-bottom:0.3rem;">External Case ID(s)</div>', unsafe_allow_html=True)

    if 'case_ids' not in st.session_state:
        st.session_state.case_ids = ['']

    case_ids_valid = []
    for i, cid in enumerate(st.session_state.case_ids):
        c1, c2 = st.columns([5, 1])
        with c1:
            val = st.text_input(
                f'Case ID {i+1}',
                value=cid,
                placeholder='e.g. 1642976#1',
                key=f'case_id_{i}',
                label_visibility='collapsed'
            )
            st.session_state.case_ids[i] = val
            if val.strip():
                case_ids_valid.append(val.strip())
        with c2:
            if i == len(st.session_state.case_ids) - 1:
                if st.button('＋', key=f'add_case_{i}', help='Add another Case ID'):
                    st.session_state.case_ids.append('')
                    st.rerun()
            else:
                if st.button('✕', key=f'rem_case_{i}', help='Remove'):
                    st.session_state.case_ids.pop(i)
                    st.rerun()

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --- COUNTRY & REGION ---
    st.markdown('<div class="section-label">02 · Country & Region</div>', unsafe_allow_html=True)

    country_options = [f"{c} ({r})" for c, r, rn in countries_list]
    country_selection = st.selectbox(
        'Country',
        options=['— Select country —'] + country_options,
        key='country_select'
    )

    selected_country = None
    selected_region = None
    if country_selection != '— Select country —':
        selected_country = country_selection.split(' (')[0]
        selected_region = get_region(selected_country, config)
        region_name = config['regions'][selected_region]['name']
        badge_class = 'region-mcc' if selected_region == 'MCC' else 'region-cs'
        st.markdown(
            f'<span class="region-badge {badge_class}">▸ {selected_region} · {region_name}</span>',
            unsafe_allow_html=True
        )

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --- DOMAIN INFORMATION ---
    st.markdown('<div class="section-label">03 · Domain Information</div>', unsafe_allow_html=True)

    primary_domain = st.text_input(
        'Primary Website / Domain',
        placeholder='e.g. esperanza.mx',
        key='primary_domain'
    )

    st.markdown('<div style="font-size:0.85rem; color:#94a3b8; margin-bottom:0.3rem;">Additional Actionable Domains</div>', unsafe_allow_html=True)

    if 'extra_domains' not in st.session_state:
        st.session_state.extra_domains = ['']

    extra_domains_valid = []
    for i, dom in enumerate(st.session_state.extra_domains):
        c1, c2 = st.columns([5, 1])
        with c1:
            val = st.text_input(
                f'Domain {i+1}',
                value=dom,
                placeholder='e.g. company.com.mx',
                key=f'extra_domain_{i}',
                label_visibility='collapsed'
            )
            st.session_state.extra_domains[i] = val
            if val.strip():
                extra_domains_valid.append(val.strip())
        with c2:
            if i == len(st.session_state.extra_domains) - 1:
                if st.button('＋', key=f'add_dom_{i}', help='Add domain'):
                    st.session_state.extra_domains.append('')
                    st.rerun()
            else:
                if st.button('✕', key=f'rem_dom_{i}', help='Remove'):
                    st.session_state.extra_domains.pop(i)
                    st.rerun()

# ===========================================================================
# RIGHT COLUMN — File Uploads
# ===========================================================================

with right_col:

    # --- MACHINE FILES ---
    st.markdown('<div class="section-label">04 · Machine Files</div>', unsafe_allow_html=True)
    machine_files = st.file_uploader(
        'Upload one or more Machine export files (.xlsx)',
        type=['xlsx'],
        accept_multiple_files=True,
        key='machine_files',
        label_visibility='collapsed'
    )
    if machine_files:
        st.markdown('<div class="tag-row">' +
            ''.join(f'<span class="tag tag-accent">📄 {f.name}</span>' for f in machine_files) +
            '</div>', unsafe_allow_html=True)

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --- CASE EVENT FILES ---
    st.markdown('<div class="section-label">05 · Case Event Files</div>', unsafe_allow_html=True)
    event_files = st.file_uploader(
        'Upload one or more Case Events export files (.xlsx)',
        type=['xlsx'],
        accept_multiple_files=True,
        key='event_files',
        label_visibility='collapsed'
    )
    if event_files:
        st.markdown('<div class="tag-row">' +
            ''.join(f'<span class="tag tag-accent">📄 {f.name}</span>' for f in event_files) +
            '</div>', unsafe_allow_html=True)

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # --- TEMPLATE FILE ---
    st.markdown('<div class="section-label">06 · Template File</div>', unsafe_allow_html=True)

    if selected_region:
        badge_class = 'region-mcc' if selected_region == 'MCC' else 'region-cs'
        st.markdown(
            f'<div class="alert alert-info">Upload the <strong>{selected_region}</strong> template for <strong>{selected_country}</strong></div>',
            unsafe_allow_html=True
        )

    template_file = st.file_uploader(
        'Upload Evidence Report Template (.xlsx)',
        type=['xlsx'],
        accept_multiple_files=False,
        key='template_file',
        label_visibility='collapsed'
    )
    if template_file:
        st.markdown(
            f'<div class="tag-row"><span class="tag tag-accent">📋 {template_file.name}</span></div>',
            unsafe_allow_html=True
        )

# ---------------------------------------------------------------------------
# VALIDATION & GENERATE
# ---------------------------------------------------------------------------

st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

# Validation checks
checks = {
    'Entity name': bool(entity_name and entity_name.strip()),
    'Case ID(s)': bool(case_ids_valid),
    'Country selected': bool(selected_country),
    'Primary domain': bool(primary_domain and primary_domain.strip()),
    'Machine file(s)': bool(machine_files),
    'Case event file(s)': bool(event_files),
    'Template file': bool(template_file),
}

all_valid = all(checks.values())

st.markdown('<div class="section-label">07 · Validation</div>', unsafe_allow_html=True)

check_cols = st.columns(4)
items = list(checks.items())
for idx, (label, ok) in enumerate(items):
    with check_cols[idx % 4]:
        dot = 'dot-green' if ok else 'dot-red'
        st.markdown(
            f'<div class="status-row"><div class="dot {dot}"></div>{label}</div>',
            unsafe_allow_html=True
        )

st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

# Generate button
gen_col, _ = st.columns([1, 2])
with gen_col:
    generate = st.button(
        '⚡ Generate Evidence Report',
        disabled=not all_valid,
        use_container_width=True,
        key='generate_btn'
    )

# ---------------------------------------------------------------------------
# PROCESSING
# ---------------------------------------------------------------------------

if generate and all_valid:
    with st.spinner('Processing files...'):
        try:
            # Load machine DataFrames
            machines_dfs = []
            for f in machine_files:
                xl = pd.ExcelFile(f)
                sheet = 'Exported Machines' if 'Exported Machines' in xl.sheet_names else xl.sheet_names[0]
                df = pd.read_excel(f, sheet_name=sheet, dtype=str)
                machines_dfs.append(df)

            # Load case event DataFrames
            events_dfs = []
            for f in event_files:
                xl = pd.ExcelFile(f)
                sheet = 'Exported Case Events' if 'Exported Case Events' in xl.sheet_names else xl.sheet_names[0]
                df = pd.read_excel(f, sheet_name=sheet, dtype={'Machine ID': str})
                events_dfs.append(df)

            # Load template workbook
            template_wb = openpyxl.load_workbook(template_file)

            # Run processing pipeline
            rows, globals_data = run_processing(
                machines_dfs=machines_dfs,
                events_dfs=events_dfs,
                primary_domain=primary_domain.strip(),
                additional_domains=extra_domains_valid,
                country=selected_country,
            )

            # Fill template
            filled_wb, template_type = fill_template(
                template_wb=template_wb,
                rows=rows,
                globals_data=globals_data,
                case_ids=case_ids_valid,
                entity_name=entity_name.strip(),
                country=selected_country,
            )

            # Save to buffer
            output_buffer = io.BytesIO()
            patch_and_save(filled_wb, output_buffer)
            output_buffer.seek(0)

            # Build filename
            output_filename = f"{entity_name.strip()} - Evidence Report.xlsx"

            # Store results in session state
            st.session_state['result_buffer'] = output_buffer.getvalue()
            st.session_state['result_filename'] = output_filename
            st.session_state['result_rows'] = rows
            st.session_state['result_globals'] = globals_data
            st.session_state['result_type'] = template_type
            st.session_state['processed'] = True

        except Exception as e:
            st.error(f'Processing error: {str(e)}')
            import traceback
            st.code(traceback.format_exc())

# ---------------------------------------------------------------------------
# RESULTS
# ---------------------------------------------------------------------------

if st.session_state.get('processed'):
    rows = st.session_state['result_rows']
    globals_data = st.session_state['result_globals']
    template_type = st.session_state['result_type']
    filename = st.session_state['result_filename']
    buffer = st.session_state['result_buffer']

    st.markdown('<div class="section-label">08 · Results</div>', unsafe_allow_html=True)

    # Metrics
    excluded_count = sum(1 for r in rows if r.get('is_excluded'))
    valid_count = len(rows) - excluded_count

    st.markdown(f"""
    <div class="result-grid">
        <div class="metric-box">
            <div class="metric-val">{globals_data['total_machines']}</div>
            <div class="metric-lbl">Valid Machines</div>
        </div>
        <div class="metric-box">
            <div class="metric-val">{globals_data['total_licenses']}</div>
            <div class="metric-lbl">Total Licenses</div>
        </div>
        <div class="metric-box">
            <div class="metric-val">{globals_data['total_events']}</div>
            <div class="metric-lbl">Valid Events</div>
        </div>
        <div class="metric-box">
            <div class="metric-val">{globals_data['years_of_use']}</div>
            <div class="metric-lbl">Years of Use</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="result-card" style="margin-top:1rem;">
        <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
                <div style="font-family:var(--mono); font-size:0.7rem; color:var(--text-muted); letter-spacing:0.1em; text-transform:uppercase;">Period</div>
                <div style="font-size:1rem; font-weight:600; margin-top:0.25rem;">{globals_data['period']}</div>
            </div>
            <div>
                <div style="font-family:var(--mono); font-size:0.7rem; color:var(--text-muted); letter-spacing:0.1em; text-transform:uppercase;">Versions</div>
                <div style="font-family:var(--mono); font-size:0.85rem; margin-top:0.25rem;">{globals_data['versions_str']}</div>
            </div>
            <div>
                <div style="font-family:var(--mono); font-size:0.7rem; color:var(--text-muted); letter-spacing:0.1em; text-transform:uppercase;">Template</div>
                <span class="region-badge {'region-mcc' if template_type == 'MCC' else 'region-cs'}">{template_type}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if excluded_count > 0:
        st.markdown(
            f'<div class="alert alert-warn">⚠ {excluded_count} machine group(s) fully excluded (Education / Commercial / Evaluation only)</div>',
            unsafe_allow_html=True
        )

    # Machine summary table
    with st.expander('View machine rows preview', expanded=False):
        preview_data = []
        for r in rows:
            preview_data.append({
                'MAC': r['active_mac'],
                'Product': r['product'],
                'Licenses': r['license_count'],
                'Version': r['version'],
                'Event Type': r['event_type'],
                'First Event': str(r['first_event']) if r['first_event'] else '-',
                'Last Event': str(r['last_event']) if r['last_event'] else '-',
                'Country': r['ip_country'],
                'Email': r['client_email'],
                'Excluded': '🔴' if r.get('is_excluded') else '✅',
            })
        st.dataframe(pd.DataFrame(preview_data), use_container_width=True, hide_index=True)

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # Download button
    dl_col, _ = st.columns([1, 2])
    with dl_col:
        st.download_button(
            label=f'⬇ Download {filename}',
            data=buffer,
            file_name=filename,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )

    st.markdown(
        '<div class="alert alert-success" style="margin-top:0.75rem;">✓ Report generated successfully. Download and verify the output before sharing.</div>',
        unsafe_allow_html=True
    )

# ---------------------------------------------------------------------------
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
