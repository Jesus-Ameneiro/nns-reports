"""
NNS Evidence Report Processor
Translates VBA macro logic to Python, handles MCC and CS templates.

Key logic:
- Each Machine ID = one report row
- Active MAC in report = most frequent MAC from case events for that Machine ID
- No cross-machine grouping by MAC
"""

import ast
import pandas as pd
from collections import defaultdict


# ---------------------------------------------------------------------------
# HELPERS
# ---------------------------------------------------------------------------

def parse_email_list(raw):
    if raw is None:
        return []
    try:
        if pd.isna(raw):
            return []
    except Exception:
        pass
    raw = str(raw).strip()
    if not raw or raw == '[]' or raw == '-':
        return []
    if raw.startswith('['):
        try:
            items = ast.literal_eval(raw)
            return [str(e).strip() for e in items if e and str(e).strip()]
        except Exception:
            raw = raw.strip("[]").replace("'", "").replace('"', '')
    return [e.strip() for e in raw.split(',') if e.strip()]


def parse_count_field(raw):
    if raw is None:
        return []
    try:
        if pd.isna(raw):
            return []
    except Exception:
        pass
    results = []
    for part in str(raw).split(','):
        part = part.strip()
        if ':' in part:
            val, _, cnt = part.rpartition(':')
            try:
                results.append((val.strip(), int(cnt.strip())))
            except ValueError:
                results.append((val.strip(), 1))
        elif part:
            results.append((part, 1))
    return results


def domain_match(email, domains):
    if not email or not domains:
        return False
    email = email.lower().strip()
    if '@' not in email:
        return False
    email_domain = email.split('@')[-1]
    for d in domains:
        d = d.lower().strip()
        if email_domain == d or email_domain.endswith('.' + d):
            return True
    return False


def select_email(client_emails, additional_emails, all_domains):
    for email in client_emails:
        if domain_match(email, all_domains):
            return email
    for email in additional_emails:
        if domain_match(email, all_domains):
            return email
    return '-'


def clean_version(v):
    """
    Normalize version to a 4-digit year string.
    Handles:
      - 4-digit years already: 2023, 2024.0  -> '2023', '2024'
      - 2-digit semver prefix: 23.1.0, 24.0  -> '2023', '2024'
      - Plain 2-digit int: 23, 24            -> '2023', '2024'
    """
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    s = str(v).strip().rstrip('.')
    if not s:
        return None
    # Use only the first numeric segment (before any dot)
    first_seg = s.split('.')[0].strip()
    try:
        n = int(first_seg)
        if 2000 <= n <= 2099:   # already a 4-digit year
            return str(n)
        if 0 <= n <= 99:        # 2-digit year prefix (e.g. 23 -> 2023)
            return str(2000 + n)
    except (ValueError, OverflowError):
        pass
    return s if s else None


def is_excluded_type(event_type):
    et = str(event_type).strip()
    for label in ['Education', 'Commercial', 'Evaluation']:
        if label.lower() in et.lower():
            return True, label
    return False, None


def compute_period(years):
    if not years:
        return '-'
    sorted_years = sorted(int(y) for y in years)
    ranges = []
    start = end = sorted_years[0]
    for y in sorted_years[1:]:
        if y == end + 1:
            end = y
        else:
            ranges.append(f"{start}-{end}" if start != end else str(start))
            start = end = y
    ranges.append(f"{start}-{end}" if start != end else str(start))
    return ', '.join(ranges)


def winning_key(d):
    if not d:
        return None
    return max(d, key=d.get)


# ---------------------------------------------------------------------------
# STEP 1: MERGE MACHINE FILES
# ---------------------------------------------------------------------------

def merge_machines(dfs):
    combined = pd.concat(dfs, ignore_index=True)
    combined['Machine ID'] = combined['Machine ID'].astype(str).str.strip()

    merged = {}
    for _, row in combined.iterrows():
        mid = row['Machine ID']
        if mid not in merged:
            merged[mid] = {col: [] for col in combined.columns if col != 'Machine ID'}
            merged[mid]['Machine ID'] = mid
        for col in combined.columns:
            if col == 'Machine ID':
                continue
            val = row[col]
            if pd.notna(val) and str(val).strip():
                merged[mid][col].append(str(val).strip())

    single_val_cols = {'Active MAC', 'Approval Status', 'Automation Status'}
    records = []
    for mid, data in merged.items():
        rec = {'Machine ID': mid}
        for col, vals in data.items():
            if col == 'Machine ID':
                continue
            unique_vals = list(dict.fromkeys(vals))
            if col in single_val_cols:
                rec[col] = unique_vals[0] if unique_vals else None
            else:
                rec[col] = ', '.join(unique_vals) if unique_vals else None
        records.append(rec)

    return pd.DataFrame(records)


# ---------------------------------------------------------------------------
# STEP 2: MERGE CASE EVENT FILES
# ---------------------------------------------------------------------------

def merge_case_events(dfs):
    combined = pd.concat(dfs, ignore_index=True)
    combined['Machine ID'] = combined['Machine ID'].astype(str).str.strip()
    dedup_cols = ['Machine ID', 'Server Timestamp', 'Event Type', 'Product']
    available = [c for c in dedup_cols if c in combined.columns]
    return combined.drop_duplicates(subset=available).reset_index(drop=True)


# ---------------------------------------------------------------------------
# STEP 3: PROCESS CASE EVENTS
# ---------------------------------------------------------------------------

def process_case_events(events_df):
    machines = {}

    for _, row in events_df.iterrows():
        mid = str(row.get('Machine ID', '')).strip()
        if not mid or mid == 'nan':
            continue

        raw_date = row.get('Server Timestamp', None)
        try:
            current_date = pd.to_datetime(raw_date)
            if pd.isna(current_date):
                current_date = None
        except Exception:
            current_date = None

        raw_product = str(row.get('Product', '') or '').strip()

        # Version column takes priority - it contains the definitive product year
        version_val = clean_version(row.get('Version', None))

        # Strip year from Product string to get the base product name
        base_product = raw_product
        if len(raw_product) >= 4 and raw_product[-4:].isdigit():
            base_product = raw_product[:-4].strip()
            # Only use Product year as fallback if Version column gave nothing
            if not version_val:
                version_val = raw_product[-4:]

        event_type  = str(row.get('Event Type', '') or '').strip()
        country     = str(row.get('Public IP Country', '') or '').strip()
        hostname    = str(row.get('Hostname', '') or '').strip()
        username    = str(row.get('Username', '') or '').strip()
        comp_domain = str(row.get('Computer Domain', '') or '').strip()

        ce_raw = row.get('Client Email Address', None) or row.get('Client Email Address.1', None)
        ae_raw = row.get('Additional Email Addresses', None) or row.get('Additional Email Addresses.1', None)
        client_emails_ev = parse_email_list(ce_raw)
        add_emails_ev    = parse_email_list(ae_raw)

        active_mac_ev = str(row.get('Active Mac', '') or '').strip()

        if mid not in machines:
            machines[mid] = {
                'base_products':      defaultdict(int),
                'valid_event_types':  defaultdict(int),
                'versions':           set(),
                'licenses':           set(),
                'machine_years':      set(),
                'countries':          defaultdict(int),
                'hostnames':          set(),
                'usernames':          set(),
                'computer_domains':   set(),
                'first_event':        None,
                'last_event':         None,
                'total_events':       0,
                'excluded_count':     0,
                'last_excluded_type': '',
                'client_emails_ev':   [],
                'add_emails_ev':      [],
                'active_macs_ev':     defaultdict(int),
            }

        m = machines[mid]
        m['total_events'] += 1

        if base_product:
            m['base_products'][base_product] += 1

        excl, excl_label = is_excluded_type(event_type)
        if excl:
            m['excluded_count'] += 1
            m['last_excluded_type'] = excl_label
        else:
            if event_type:
                m['valid_event_types'][event_type] += 1
            if version_val:
                m['versions'].add(version_val)
                if current_date:
                    year = current_date.year
                    m['machine_years'].add(year)
                    m['licenses'].add(f"{version_val}|{year}")

        if current_date:
            if m['first_event'] is None or current_date < m['first_event']:
                m['first_event'] = current_date
            if m['last_event'] is None or current_date > m['last_event']:
                m['last_event'] = current_date

        if hostname and '=' not in hostname and len(hostname) < 50:
            m['hostnames'].add(hostname)
        if username and '=' not in username and len(username) < 50:
            m['usernames'].add(username)
        if comp_domain and comp_domain != 'nan' and '=' not in comp_domain and len(comp_domain) < 100:
            m['computer_domains'].add(comp_domain)
        if country:
            m['countries'][country] += 1

        for e in client_emails_ev:
            if '=' not in e and len(e) < 100 and e not in m['client_emails_ev']:
                m['client_emails_ev'].append(e)
        for e in add_emails_ev:
            if '=' not in e and len(e) < 100 and e not in m['add_emails_ev']:
                m['add_emails_ev'].append(e)

        if active_mac_ev and len(active_mac_ev) == 17:
            m['active_macs_ev'][active_mac_ev] += 1

    return machines


# ---------------------------------------------------------------------------
# STEP 4: ENRICH WITH MACHINE FILE DATA
# ---------------------------------------------------------------------------

def enrich_with_machines(machines, machines_df, all_domains):
    machines_df['Machine ID'] = machines_df['Machine ID'].astype(str).str.strip()
    machine_lookup = {str(r['Machine ID']).strip(): r for _, r in machines_df.iterrows()}

    for mid, m in machines.items():
        row = machine_lookup.get(mid)

        # active_mac: most frequent from case events; fallback to machines file
        if m['active_macs_ev']:
            m['active_mac'] = winning_key(m['active_macs_ev'])
        elif row is not None:
            mf_mac = str(row.get('Active MAC', '') or '').strip()
            m['active_mac'] = mf_mac if mf_mac and mf_mac != 'nan' else '-'
        else:
            m['active_mac'] = '-'

        # Computer domains from machines file
        if row is not None:
            cd_raw = str(row.get('Computer Domains', '') or '').strip()
            if cd_raw and cd_raw != 'nan':
                for val, _ in parse_count_field(cd_raw):
                    if val and '=' not in val:
                        m['computer_domains'].add(val)

        # Filter computer_domains to only those matching user-provided domains.
        # A computer domain matches if it equals or is a sub-domain of any
        # entry in all_domains (case-insensitive).  Unrecognised domains such
        # as 'coinsa.com.ar' are silently dropped.
        def _domain_matches(cd, domains):
            cd = cd.lower().strip()
            for d in domains:
                d = d.lower().strip()
                if cd == d or cd.endswith('.' + d):
                    return True
            return False

        if all_domains:
            m['computer_domains'] = {
                cd for cd in m['computer_domains']
                if _domain_matches(cd, all_domains)
            }

        # Emails: machines file priority, events as supplement
        client_emails_mf = []
        add_emails_mf    = []
        if row is not None:
            client_emails_mf = parse_email_list(row.get('Client Email Addresses', None))
            add_emails_mf    = parse_email_list(row.get('Additional Email Addresses', None))

        all_client     = client_emails_mf + [e for e in m['client_emails_ev'] if e not in client_emails_mf]
        all_additional = add_emails_mf    + [e for e in m['add_emails_ev']    if e not in add_emails_mf]

        m['selected_email'] = select_email(all_client, all_additional, all_domains)

    return machines


# ---------------------------------------------------------------------------
# STEP 5: BUILD FINAL ROW DATA
# ---------------------------------------------------------------------------

def build_rows(machines):
    rows = []
    global_years    = set()
    global_versions = set()
    global_licenses = 0
    valid_events    = 0
    total_valid     = 0

    for mid in sorted(machines.keys()):
        m = machines[mid]
        is_excluded = (m['total_events'] > 0 and m['total_events'] == m['excluded_count'])

        wp = winning_key(m['base_products']) or '-'
        if wp and len(wp) >= 4 and wp[-4:].isdigit():
            wp = wp[:-4].strip()

        # Normalize: the report always shows "SketchUp Pro" regardless of the
        # source product name (SketchUp Make, SketchUp, SketchUp Pro, etc.).
        if not is_excluded and wp not in ('-', 'N/A'):
            wp = 'SketchUp Pro'

        # Normalize: non-excluded machines always show "Unlicensed" in the
        # report regardless of the actual event type (Personal, Undefined, etc.).
        wt = (m['last_excluded_type'] or 'Excluded') if is_excluded else 'Unlicensed'
        wc = winning_key(m['countries']) or '-'

        sorted_vers = sorted(m['versions'], key=lambda x: (len(x), x))
        version_str  = ', '.join(sorted_vers) if sorted_vers else '-'
        hostnames    = ', '.join(sorted(m['hostnames']))          if m['hostnames']          else '-'
        usernames    = ', '.join(sorted(m['usernames']))          if m['usernames']          else '-'
        comp_domains = ', '.join(sorted(m['computer_domains']))   if m['computer_domains']   else '-'

        license_count = len(m['licenses']) if not is_excluded else 'N/A'
        first_ev = m['first_event'].date() if m['first_event'] else None
        last_ev  = m['last_event'].date()  if m['last_event']  else None

        # Skip machines whose every event was Commercial, Education, or Evaluation.
        # These must not appear in the final report at all.
        if is_excluded:
            continue

        # Guard: skip rows with no meaningful data (corrupt timestamps / versions).
        has_data = (
            first_ev is not None
            or (isinstance(license_count, int) and license_count > 0)
            or (version_str not in ('-', 'N/A', '', None))
        )
        if not has_data:
            continue

        rows.append({
            'active_mac':      m.get('active_mac', '-'),
            'license_count':   license_count,
            'product':         wp,
            'first_event':     first_ev,
            'last_event':      last_ev,
            'event_type':      wt,
            'version':         'N/A' if is_excluded else version_str,
            'ip_country':      wc,
            'hostname':        hostnames,
            'username':        usernames,
            'client_email':    m.get('selected_email', '-'),
            'computer_domain': comp_domains,
            'is_excluded':     is_excluded,
        })

        if not is_excluded:
            total_valid += 1
            global_years.update(m['machine_years'])
            global_versions.update(m['versions'])
            if isinstance(license_count, int):
                global_licenses += license_count
            valid_events += (m['total_events'] - m['excluded_count'])

    sorted_gv = sorted(global_versions, key=lambda x: (len(x), x))
    globals_data = {
        'total_machines': total_valid,
        'total_users':    total_valid,
        'versions_str':   ', '.join(sorted_gv) if sorted_gv else '-',
        'total_versions': len(sorted_gv),
        'total_events':   valid_events,
        'total_licenses': global_licenses,
        'years_of_use':   len(global_years),
        'period':         compute_period(global_years),
        'country':        '',
    }

    return rows, globals_data


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def run_processing(machines_dfs, events_dfs, primary_domain, additional_domains):
    all_domains = [d.strip() for d in [primary_domain] + additional_domains if d.strip()]
    machines_df = merge_machines(machines_dfs)
    events_df   = merge_case_events(events_dfs)
    machines    = process_case_events(events_df)
    machines    = enrich_with_machines(machines, machines_df, all_domains)
    return build_rows(machines)
