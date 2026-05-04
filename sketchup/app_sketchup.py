"""
SketchUp (Trimble) — NNS Evidence Report Generator
MCC and CS regions.
All business logic lives in sketchup/processor.py and sketchup/report_writer.py.
This module exposes a single function:  render()
Called by the root app.py inside the SketchUp tab.
"""

import base64
import io
import json
import openpyxl
import pandas as pd
import streamlit as st
import traceback
from pathlib import Path

from sketchup.processor    import run_processing
from sketchup.report_writer import fill_template, patch_and_save

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

CONFIG_PATH    = Path(__file__).parent / 'config.json'
TEMPLATES_DIR  = Path(__file__).parent.parent / 'templates'

# Template filenames inside templates/
_TMPL_MCC     = 'MCC_-_Evidence_Report.xlsx'
_TMPL_CS_ARG  = 'CS_-_Evidence_Report_SOUTH_(ARGENTINA).xlsx'
_TMPL_CS_PAR  = 'CS_-_Evidence_Report_SOUTH_(PARAGUAY).xlsx'
_TMPL_CS_Q1   = 'CS_-_Evidence_Report_SOUTH.xlsx'   # all other CS countries

# Countries with dedicated CS templates
_CS_COUNTRY_TEMPLATES = {
    'Argentina': _TMPL_CS_ARG,
    'Paraguay':  _TMPL_CS_PAR,
}


def _load_config():
    with open(CONFIG_PATH) as f:
        return json.load(f)


def _get_region(country, config):
    for region, data in config['regions'].items():
        if country in data['countries']:
            return region
    return None


def _all_countries(config):
    countries = []
    for region, data in config['regions'].items():
        countries.extend(data['countries'])
    return sorted(countries)


def _get_template_path(country, config):
    """
    Return the Path to the correct pre-stored Evidence Report template
    based on the selected country.

    MCC  → templates/MCC_-_Evidence_Report.xlsx
    CS:
      Argentina → templates/CS_-_Evidence_Report_SOUTH__ARGENTINA_.xlsx
      Paraguay  → templates/CS_-_Evidence_Report_SOUTH__PARAGUAY_.xlsx
      Others    → templates/CS_-_Evidence_Report_SOUTH.xlsx
    """
    region = _get_region(country, config)
    if region == 'MCC':
        return TEMPLATES_DIR / _TMPL_MCC
    # CS
    fname = _CS_COUNTRY_TEMPLATES.get(country, _TMPL_CS_Q1)
    return TEMPLATES_DIR / fname


def _tip(text: str) -> str:
    safe = text.replace("'", "&#39;").replace('"', '&quot;')
    return (
        f'<span class="tip-wrap">'
        f'<span class="tip-badge">?</span>'
        f'<span class="tip-box">{safe}</span>'
        f'</span>'
    )


# ---------------------------------------------------------------------------
# TOOLTIP CSS
# ---------------------------------------------------------------------------

_TOOLTIP_CSS = """
<style>
.tip-wrap {
    position: relative;
    display: inline-flex;
    align-items: center;
    margin-left: 0.4rem;
    vertical-align: middle;
}
.tip-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 16px;
    height: 16px;
    border-radius: 50%;
    background: rgba(249,115,22,0.15);
    border: 1px solid rgba(249,115,22,0.5);
    color: #f97316;
    font-family: 'DM Mono', monospace;
    font-size: 0.6rem;
    font-weight: 700;
    cursor: default;
    user-select: none;
    flex-shrink: 0;
}
.tip-box {
    visibility: hidden;
    opacity: 0;
    position: absolute;
    left: 22px;
    top: 50%;
    transform: translateY(-50%);
    background: #1a1a1a;
    border: 1px solid rgba(249,115,22,0.35);
    border-radius: 6px;
    padding: 0.5rem 0.75rem;
    width: 240px;
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #d4d4d4;
    line-height: 1.5;
    z-index: 9999;
    pointer-events: none;
    transition: opacity 0.15s ease;
    white-space: normal;
    box-shadow: 0 4px 20px rgba(0,0,0,0.5);
}
.tip-wrap:hover .tip-box {
    visibility: visible;
    opacity: 1;
}
</style>
"""

# ---------------------------------------------------------------------------
# LABEL HELPERS
# ---------------------------------------------------------------------------

def _label(text: str, tip: str = '', style: str = '') -> None:
    style_attr = f' style="{style}"' if style else ''
    tip_html   = _tip(tip) if tip else ''
    st.markdown(
        f'<div class="section-label"{style_attr}>{text}{tip_html}</div>',
        unsafe_allow_html=True,
    )


def _sublabel(text: str, tip: str = '') -> None:
    tip_html = _tip(tip) if tip else ''
    st.markdown(
        f'<div class="sub-label">{text}{tip_html}</div>',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# USER MANUAL BUTTON
# ---------------------------------------------------------------------------

@st.cache_data
def _manual_b64() -> str:
    """
    Read the PDF and return a base64-encoded data URI.
    Cached so the file is only read once per session.
    Using a data URI instead of /static/... avoids Streamlit Community
    Cloud's /app/static/ path ambiguity and works without any
    static file serving config.
    """
    manual_path = Path(__file__).parent.parent / 'docs' / 'user_manual.pdf'
    if not manual_path.exists():
        return ''
    data = manual_path.read_bytes()
    return base64.b64encode(data).decode()


def _manual_button() -> None:
    b64 = _manual_b64()
    if not b64:
        return
    data_uri = f'data:application/pdf;base64,{b64}'
    st.markdown(
        f'''<a href="{data_uri}" target="_blank" style="
            display:inline-flex; align-items:center; gap:0.4rem;
            font-family:'DM Mono', monospace; font-size:0.75rem;
            color:var(--text-muted); text-decoration:none;
            padding:0.3rem 0.7rem;
            border:1px solid var(--border);
            border-radius:6px;
            background:var(--surface2);
            transition:all 0.15s;
        "
        onmouseover="this.style.borderColor='#f97316';this.style.color='#f97316';"
        onmouseout="this.style.borderColor='var(--border)';this.style.color='var(--text-muted)';">
            📖 User Manual
        </a>''',
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# MAIN RENDER FUNCTION
# ---------------------------------------------------------------------------

def render():
    """Render the full SketchUp Evidence Report Generator UI inside its tab."""

    # ── Counter-based reset ─────────────────────────────────────────────────
    _count = st.session_state.get('sk_clear_count', 0)

    def _clear():
        c = st.session_state.get('sk_clear_count', 0)
        for k in [k for k in list(st.session_state.keys())
                  if k.startswith('sk_') and k != 'sk_clear_count']:
            del st.session_state[k]
        st.session_state['sk_clear_count'] = c + 1

    st.markdown(_TOOLTIP_CSS, unsafe_allow_html=True)

    _, manual_col = st.columns([5, 1])
    with manual_col:
        _manual_button()

    config         = _load_config()
    countries_list = _all_countries(config)

    # -----------------------------------------------------------------------
    # LAYOUT
    # -----------------------------------------------------------------------

    left_col, right_col = st.columns([1, 1], gap='large')

    # =======================================================================
    # LEFT COLUMN
    # =======================================================================

    with left_col:

        # --- 01 · Case Information ---
        _label('01 · Case Information')

        entity_name = st.text_input(
            'Entity / Organization Name',
            placeholder='e.g. Acme Corp S.A.',
            key=f'sk_entity_name_{_count}',
            help='Full legal name of the company or organization being investigated.',
        )

        _cids_key = f'sk_case_ids_{_count}'
        if _cids_key not in st.session_state:
            st.session_state[_cids_key] = ['']

        _sublabel(
            'Case ID(s)',
            tip='The Pleteo case identifier(s) for this investigation, e.g. 1234567#1. '
                'Add one ID per line. Use ＋ to add more.',
        )
        for i, cid in enumerate(st.session_state[_cids_key]):
            c1, c2, c3 = st.columns([6, 1, 1])
            with c1:
                st.session_state[_cids_key][i] = st.text_input(
                    f'Case ID {i+1}', value=cid,
                    label_visibility='collapsed',
                    placeholder='e.g. 1234567#1',
                    key=f'sk_cid_{i}_{_count}',
                )
            with c2:
                if st.button('＋', key=f'sk_add_case_{i}_{_count}',
                             help='Add another Case ID'):
                    st.session_state[_cids_key].append('')
                    st.rerun()
            with c3:
                if len(st.session_state[_cids_key]) > 1:
                    if st.button('✕', key=f'sk_rem_case_{i}_{_count}',
                                 help='Remove'):
                        st.session_state[_cids_key].pop(i)
                        st.rerun()

        case_ids_valid = [c.strip() for c in st.session_state[_cids_key]
                          if c.strip()]

        # --- 02 · Country & Region ---
        _label(
            '02 · Country & Region',
            tip='Country where the entity operates. Determines the report region '
                '(MCC or Cono Sur) and selects the correct Evidence Report template automatically.',
            style='margin-top:1.5rem;',
        )

        selected_country = st.selectbox(
            'Country',
            options=[''] + countries_list,
            key=f'sk_country_{_count}',
        )

        selected_region = None
        if selected_country:
            selected_region = _get_region(selected_country, config)
            region_name     = config['regions'][selected_region]['name']
            badge_class     = 'region-mcc' if selected_region == 'MCC' else 'region-cs'

            # Show region badge + which template will be used
            tmpl_path = _get_template_path(selected_country, config)
            st.markdown(
                f'<span class="region-badge {badge_class}">'
                f'{selected_region} · {region_name}</span>'
                f'&nbsp;&nbsp;'
                f'<span style="font-family:var(--mono);font-size:0.72rem;'
                f'color:var(--text-muted);">📋 {tmpl_path.name}</span>',
                unsafe_allow_html=True,
            )

        # --- 03 · Domain Information ---
        _label(
            '03 · Domain Information',
            tip='Add every domain associated with this entity: '
                'email domain, web domain, computer/AD domain, and any subsidiary domains. '
                'These are used to filter and match machines, emails, and computer domains in the report. '
                'Enter the primary domain below, then press ＋ to add each additional domain.',
            style='margin-top:1.5rem;',
        )

        primary_domain = st.text_input(
            'Primary Domain',
            placeholder='e.g. company.com',
            key=f'sk_primary_domain_{_count}',
            help='The main domain of the entity (e.g. company.com). '
                 'This is used to match email addresses, computer domains, '
                 'and filter out unrelated machines.',
        )

        _doms_key = f'sk_extra_domains_{_count}'
        if _doms_key not in st.session_state:
            st.session_state[_doms_key] = []

        if st.session_state[_doms_key]:
            _sublabel(
                'Additional Domains',
                tip='Any other domain belonging to this entity: subsidiaries, '
                    'regional offices, email aliases, or Active Directory domains. '
                    'Press ＋ to add as many as needed.',
            )
        for i, dom in enumerate(st.session_state[_doms_key]):
            c1, c2 = st.columns([7, 1])
            with c1:
                st.session_state[_doms_key][i] = st.text_input(
                    f'Domain {i+1}', value=dom,
                    label_visibility='collapsed',
                    placeholder='e.g. subsidiary.com',
                    key=f'sk_dom_{i}_{_count}',
                )
            with c2:
                if st.button('✕', key=f'sk_rem_dom_{i}_{_count}', help='Remove'):
                    st.session_state[_doms_key].pop(i)
                    st.rerun()

        if st.button('＋ Add domain', key=f'sk_add_dom_{_count}'):
            st.session_state[_doms_key].append('')
            st.rerun()

        extra_domains_valid = [d.strip() for d in st.session_state[_doms_key]
                               if d.strip()]

    # =======================================================================
    # RIGHT COLUMN
    # =======================================================================

    with right_col:

        # --- 04 · Machine Files ---
        _label(
            '04 · Machine Files',
            tip='Export from Pleteo: the "Exported Machines" sheet. '
                'One or more .xlsx files are accepted — upload all files '
                'belonging to this case.',
        )
        machine_files = st.file_uploader(
            'Upload exported machine file(s)',
            type=['xlsx'],
            accept_multiple_files=True,
            key=f'sk_machine_files_{_count}',
        )
        if machine_files:
            for f in machine_files:
                st.markdown(
                    f'<div class="file-tag">📄 {f.name}</div>',
                    unsafe_allow_html=True,
                )

        # --- 05 · Case Event Files ---
        _label(
            '05 · Case Event Files',
            tip='Export from Pleteo: the "Exported Case Events" sheet. '
                'One or more .xlsx files are accepted — upload all event files '
                'for this case. Must correspond to the same machines uploaded above.',
            style='margin-top:1.5rem;',
        )
        event_files = st.file_uploader(
            'Upload exported case event file(s)',
            type=['xlsx'],
            accept_multiple_files=True,
            key=f'sk_event_files_{_count}',
        )
        if event_files:
            for f in event_files:
                st.markdown(
                    f'<div class="file-tag">📄 {f.name}</div>',
                    unsafe_allow_html=True,
                )

    # -----------------------------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------------------------

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    checks = {
        'Entity name':     bool(entity_name and entity_name.strip()),
        'Case ID(s)':      bool(case_ids_valid),
        'Country':         bool(selected_country),
        'Primary domain':  bool(primary_domain and primary_domain.strip()),
        'Machine file(s)': bool(machine_files),
        'Case event(s)':   bool(event_files),
    }
    all_valid = all(checks.values())

    st.markdown('<div class="section-label">06 · Validation</div>',
                unsafe_allow_html=True)
    check_cols = st.columns(3)
    for idx, (lbl, ok) in enumerate(checks.items()):
        with check_cols[idx % 3]:
            dot = 'dot-green' if ok else 'dot-red'
            st.markdown(
                f'<div class="status-row">'
                f'<div class="dot {dot}"></div>{lbl}</div>',
                unsafe_allow_html=True,
            )

    st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

    # -----------------------------------------------------------------------
    # GENERATE + CLEAR
    # -----------------------------------------------------------------------

    gen_col, clear_col, _ = st.columns([2, 1, 2])
    with gen_col:
        generate = st.button(
            '⚡ Generate Evidence Report',
            disabled=not all_valid,
            use_container_width=True,
            key=f'sk_generate_btn_{_count}',
        )
    with clear_col:
        if st.button(
            '🗑 Clear Data',
            use_container_width=True,
            key='sk_clear_btn',
            type='secondary',
            help='Clear all inputs, files and results — same as a page refresh',
        ):
            _clear()
            st.rerun()

    # -----------------------------------------------------------------------
    # PROCESSING
    # -----------------------------------------------------------------------

    _processed_key = f'sk_processed_{_count}'

    if generate and all_valid:
        with st.spinner('Processing files...'):
            try:
                machines_dfs, events_dfs = [], []
                for f in machine_files:
                    xl = pd.ExcelFile(f)
                    sheet = 'Exported Machines' if 'Exported Machines' in xl.sheet_names \
                            else xl.sheet_names[0]
                    machines_dfs.append(
                        pd.read_excel(f, sheet_name=sheet, dtype={'Machine ID': str})
                    )
                for f in event_files:
                    xl = pd.ExcelFile(f)
                    sheet = 'Exported Case Events' if 'Exported Case Events' in xl.sheet_names \
                            else xl.sheet_names[0]
                    events_dfs.append(
                        pd.read_excel(f, sheet_name=sheet, dtype={'Machine ID': str})
                    )

                rows, globals_data = run_processing(
                    machines_dfs=machines_dfs,
                    events_dfs=events_dfs,
                    primary_domain=primary_domain.strip(),
                    additional_domains=extra_domains_valid,
                    country=selected_country,
                )

                # ── Load template from repo — selected automatically by country ──
                tmpl_path  = _get_template_path(selected_country, config)
                _tmpl_raw  = tmpl_path.read_bytes()
                template_wb = openpyxl.load_workbook(
                    io.BytesIO(_tmpl_raw), keep_links=True)

                filled_wb, template_type = fill_template(
                    template_wb, rows, globals_data,
                    case_ids_valid, entity_name.strip(), selected_country,
                    raw_bytes=_tmpl_raw,
                )

                safe_entity = ''.join(
                    c for c in entity_name.strip()
                    if c.isalnum() or c in ' ._-'
                ).strip()
                filename = f'{safe_entity} - Evidence Report.xlsx'

                buf = io.BytesIO()
                patch_and_save(filled_wb, buf)
                buf.seek(0)

                st.session_state[_processed_key]              = True
                st.session_state[f'sk_result_rows_{_count}']    = rows
                st.session_state[f'sk_result_globals_{_count}'] = globals_data
                st.session_state[f'sk_result_type_{_count}']    = template_type
                st.session_state[f'sk_result_filename_{_count}'] = filename
                st.session_state[f'sk_result_buffer_{_count}']  = buf.read()

                st.markdown(
                    '<div class="alert alert-success">✓ Processing complete.</div>',
                    unsafe_allow_html=True,
                )

            except Exception as e:
                st.markdown(
                    f'<div class="alert alert-error">✗ Error: {e}</div>',
                    unsafe_allow_html=True,
                )
                st.text(traceback.format_exc())

    # -----------------------------------------------------------------------
    # RESULTS
    # -----------------------------------------------------------------------

    if st.session_state.get(_processed_key):
        rows         = st.session_state[f'sk_result_rows_{_count}']
        globals_data = st.session_state[f'sk_result_globals_{_count}']
        template_type = st.session_state[f'sk_result_type_{_count}']
        filename     = st.session_state[f'sk_result_filename_{_count}']
        buffer       = st.session_state[f'sk_result_buffer_{_count}']

        st.markdown('<div class="section-label">07 · Results</div>',
                    unsafe_allow_html=True)

        excluded_count = sum(1 for r in rows if r.get('is_excluded'))

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
                f'<div class="alert alert-warn">⚠ {excluded_count} machine group(s) '
                f'fully excluded (Education / Commercial / Evaluation only)</div>',
                unsafe_allow_html=True,
            )

        with st.expander('View machine rows preview', expanded=False):
            preview_data = [{
                'MAC':        r['active_mac'],
                'Product':    r['product'],
                'Licenses':   r['license_count'],
                'Version':    r['version'],
                'Event Type': r['event_type'],
                'First Event': str(r['first_event']) if r['first_event'] else '-',
                'Last Event':  str(r['last_event'])  if r['last_event']  else '-',
                'Country':    r['ip_country'],
                'Email':      r['client_email'],
                'Excluded':   '🔴' if r.get('is_excluded') else '✅',
            } for r in rows]
            st.dataframe(pd.DataFrame(preview_data),
                         use_container_width=True, hide_index=True)

        st.markdown('<div class="hr"></div>', unsafe_allow_html=True)

        dl_col, _ = st.columns([1, 2])
        with dl_col:
            st.download_button(
                label=f'⬇ Download {filename}',
                data=buffer,
                file_name=filename,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                use_container_width=True,
                key=f'sk_download_btn_{_count}',
            )

        st.markdown(
            '<div class="alert alert-success" style="margin-top:0.75rem;">'
            '✓ Report generated successfully. Download and verify the output '
            'before sharing.</div>',
            unsafe_allow_html=True,
        )
