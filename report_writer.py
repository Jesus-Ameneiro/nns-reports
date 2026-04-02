"""
NNS Report Writer - fills MCC and CS templates with processed data.
"""

import openpyxl
from openpyxl.utils import get_column_letter
from datetime import date


# ---------------------------------------------------------------------------
# UTILITIES
# ---------------------------------------------------------------------------

def detect_template_type(wb):
    return 'MCC' if 'LC Summary' in wb.sheetnames else 'CS'

def detect_summary_sheet(wb):
    for name in ['LC Summary', 'Summary', 'New Template']:
        if name in wb.sheetnames:
            return name
    return wb.sheetnames[0]

def find_col_by_header(ws, header_row, header_name):
    for cell in ws[header_row]:
        if cell.value and str(cell.value).strip().lower() == header_name.strip().lower():
            return cell.column
    return None

def safe_set(ws, row, col, value):
    """Set cell value safely, skipping merged cells that are not top-left."""
    cell = ws.cell(row=row, column=col)
    try:
        cell.value = value
    except AttributeError:
        pass  # merged cell slave - skip

def col_all_dash(ws, col_idx, data_start_row, data_end_row):
    for r in range(data_start_row, data_end_row + 1):
        v = ws.cell(row=r, column=col_idx).value
        if v is not None and str(v).strip() not in ('', '-'):
            return False
    return True

def format_date(d):
    if d is None:
        return '-'
    if hasattr(d, 'strftime'):
        return d.strftime('%Y-%m-%d')
    return str(d)

def clear_row_range(ws, row, col_start, col_end):
    """Clear cells in a row range, safely skipping merged cell slaves."""
    for c in range(col_start, col_end + 1):
        safe_set(ws, row, c, None)


# ---------------------------------------------------------------------------
# MCC TEMPLATE FILLER
# ---------------------------------------------------------------------------

def fill_mcc(wb, rows, globals_data, case_ids, entity_name, country):
    ws_summary = wb['LC Summary']
    ws_data = wb['Data']

    # LC Summary
    ws_summary['B8'] = ', '.join(case_ids)
    ws_summary['B9'] = entity_name
    ws_summary['A14'] = country
    ws_summary['B14'] = globals_data['total_machines']
    ws_summary['C14'] = globals_data['total_users']
    ws_summary['D14'] = globals_data['versions_str']
    ws_summary['E14'] = globals_data['total_events']
    ws_summary['F14'] = globals_data['total_licenses']
    ws_summary['G14'] = globals_data['years_of_use']
    ws_summary['H14'] = globals_data['period']
    ws_summary['B16'] = globals_data['total_machines']
    ws_summary['C16'] = globals_data['total_users']
    ws_summary['D16'] = globals_data['versions_str']
    ws_summary['E16'] = globals_data['total_events']
    ws_summary['F16'] = globals_data['total_licenses']
    ws_summary['G16'] = globals_data['years_of_use']

    SUMMARY_HEADER_ROW = 13
    comp_domain_col_summary = find_col_by_header(ws_summary, SUMMARY_HEADER_ROW, 'COMPUTER DOMAIN')

    # Data sheet
    DATA_HEADER_ROW = 13
    DATA_START_ROW = 14

    col_map = {}
    for cell in ws_data[DATA_HEADER_ROW]:
        if cell.value:
            col_map[str(cell.value).strip()] = cell.column

    mcc_col_order = [
        ('Active MAC',             'active_mac'),
        ('# Licenses',             'license_count'),
        ('Products',               'product'),
        ('First Event',            'first_event'),
        ('Last Event',             'last_event'),
        ('Event Types',            'event_type'),
        ('Computer Domains',       'computer_domain'),
        ('Version',                'version'),
        ('IP Country',             'ip_country'),
        ('Hostname',               'hostname'),
        ('Username',               'username'),
        ('Client Email Addresses', 'client_email'),
    ]

    # Template has 18 pre-bordered data rows (rows 14-31).
    # Write data into them, then delete any excess bordered rows.
    TEMPLATE_DATA_ROWS = 18   # rows 14..31 in the blank template
    n_rows = len(rows)

    # Write data into the first n_rows bordered slots
    for idx, row in enumerate(rows):
        r = DATA_START_ROW + idx
        for header, field in mcc_col_order:
            col_idx = col_map.get(header)
            if col_idx is None:
                continue
            val = row.get(field, '-')
            if val is None:
                val = '-'
            if field in ('first_event', 'last_event'):
                val = format_date(val)
            safe_set(ws_data, r, col_idx, val)

    data_end_row = DATA_START_ROW + n_rows - 1

    # Delete excess pre-bordered rows (from bottom up to preserve row indices)
    excess = TEMPLATE_DATA_ROWS - n_rows
    if excess > 0:
        first_excess = DATA_START_ROW + n_rows
        last_excess  = DATA_START_ROW + TEMPLATE_DATA_ROWS - 1
        ws_data.delete_rows(first_excess, excess)

    # Footer: note row is always 4 rows after the last data row
    # (gap rows: data_end+1, data_end+2, data_end+3 are blank, footer note at data_end+4)
    footer_note_row   = data_end_row + 4
    specialist_row    = footer_note_row + 3

    # Clear any stale footer content left by the template in those areas
    for r in range(data_end_row + 1, data_end_row + 15):
        for c in range(1, ws_data.max_column + 1):
            safe_set(ws_data, r, c, None)

    footer_content = [
        (footer_note_row,  'Nota: El presente documento contiene información confidencial y se proporciona exclusivamente dentro del marco de License Compliance.'),
        (specialist_row,   'XXXXXXXX'),
        (specialist_row+1, 'Especialista en Resolución'),
        (specialist_row+2, '=B2468(XXX) xxxx - xxxx'),
        (specialist_row+3, 'XXXXX@ruvixx.com'),
        (specialist_row+5, '425 Page Mill Rd, Suite 200, Palo Alto, 94306'),
    ]
    for r, text in footer_content:
        safe_set(ws_data, r, 1, text)

    # Column deletion: Computer Domains and Client Email Addresses
    cols_to_delete = []
    for header in ['Computer Domains', 'Client Email Addresses']:
        col_idx = col_map.get(header)
        if col_idx and col_all_dash(ws_data, col_idx, DATA_START_ROW, data_end_row):
            cols_to_delete.append(col_idx)

    # If Computer Domains being deleted, also remove from summary
    if col_map.get('Computer Domains') in cols_to_delete and comp_domain_col_summary:
        ws_summary.delete_cols(comp_domain_col_summary)

    for col_idx in sorted(cols_to_delete, reverse=True):
        ws_data.delete_cols(col_idx)

    return wb


# ---------------------------------------------------------------------------
# CS TEMPLATE FILLER
# ---------------------------------------------------------------------------

def fill_cs(wb, rows, globals_data, case_ids, entity_name, country):
    summary_name = detect_summary_sheet(wb)
    ws_summary = wb[summary_name]
    ws_data = wb['Data']

    # Summary sheet
    ws_summary['B9']  = ', '.join(case_ids)
    ws_summary['B10'] = country
    ws_summary['B11'] = entity_name

    # Computer Domain row (B12)
    all_comp_domains = set()
    for row in rows:
        cd = row.get('computer_domain', '-')
        if cd and cd != '-':
            for d in cd.split(','):
                d = d.strip()
                if d:
                    all_comp_domains.add(d)
    ws_summary['B12'] = ', '.join(sorted(all_comp_domains)) if all_comp_domains else '-'

    ws_summary['B13'] = globals_data['versions_str']
    ws_summary['G13'] = globals_data['total_versions']
    ws_summary['B14'] = globals_data['years_of_use']
    ws_summary['D14'] = globals_data['period']

    # Find Machines/Users/Events value row (one row above the labels row)
    machines_label_row = None
    for r in range(15, 30):
        for c in range(1, 8):
            v = ws_summary.cell(row=r, column=c).value
            if v and str(v).strip() == 'Machines':
                machines_label_row = r
                break
        if machines_label_row:
            break

    if machines_label_row:
        val_row = machines_label_row - 1
        safe_set(ws_summary, val_row, 2, globals_data['total_machines'])
        safe_set(ws_summary, val_row, 4, globals_data['total_users'])
        safe_set(ws_summary, val_row, 7, globals_data['total_events'])

    # Licensed copies for price formula
    ws_summary['C31'] = globals_data['total_licenses']

    # Data sheet header info
    safe_set(ws_data, 6, 2, ', '.join(case_ids))
    safe_set(ws_data, 7, 2, entity_name)
    safe_set(ws_data, 8, 2, country)
    safe_set(ws_data, 7, 11, '=TODAY()')

    DATA_HEADER_ROW = 11
    DATA_START_ROW  = 12

    col_map = {}
    for cell in ws_data[DATA_HEADER_ROW]:
        if cell.value:
            col_map[str(cell.value).strip()] = cell.column

    cs_col_order = [
        ('Products',               'product'),
        ('Version',                'version'),
        ('Event Types',            'event_type'),
        ('Active MAC',             'active_mac'),
        ('# Licenses',             'license_count'),
        ('First Event',            'first_event'),
        ('Last Event',             'last_event'),
        ('Computer Domains',       'computer_domain'),
        ('IP Country',             'ip_country'),
        ('Hostname',               'hostname'),
        ('Username',               'username'),
        ('Client Email Addresses', 'client_email'),
    ]

    # CS template has 12 pre-bordered data rows (rows 12-23).
    TEMPLATE_DATA_ROWS = 12
    n_rows = len(rows)

    # Write data into the first n_rows bordered slots
    for idx, row in enumerate(rows):
        r = DATA_START_ROW + idx
        for header, field in cs_col_order:
            col_idx = col_map.get(header)
            if col_idx is None:
                continue
            val = row.get(field, '-')
            if val is None:
                val = '-'
            if field in ('first_event', 'last_event'):
                val = format_date(val)
            safe_set(ws_data, r, col_idx, val)

    data_end_row = DATA_START_ROW + n_rows - 1

    # Delete excess pre-bordered rows (from bottom up)
    excess = TEMPLATE_DATA_ROWS - n_rows
    if excess > 0:
        ws_data.delete_rows(DATA_START_ROW + n_rows, excess)

    # Footer: 2 rows after last data row
    footer_note_row = data_end_row + 2

    # Clear any stale footer content
    for r in range(data_end_row + 1, data_end_row + 10):
        for c in range(1, ws_data.max_column + 1):
            safe_set(ws_data, r, c, None)

    safe_set(ws_data, footer_note_row, 1,
        'Note: This document contains confidential information and is provided '
        'exclusively within the framework of License Compliance.')

    # Column deletion
    cols_to_delete = []
    for header in ['Computer Domains', 'Client Email Addresses']:
        col_idx = col_map.get(header)
        if col_idx and col_all_dash(ws_data, col_idx, DATA_START_ROW, data_end_row):
            cols_to_delete.append(col_idx)

    # If Computer Domains deleted, also clear summary row 12
    if col_map.get('Computer Domains') in cols_to_delete:
        ws_summary['B12'] = None
        for r in range(1, ws_summary.max_row + 1):
            if ws_summary.cell(row=r, column=1).value == 'Computer Domain':
                safe_set(ws_summary, r, 1, None)
                safe_set(ws_summary, r, 2, None)
                break

    for col_idx in sorted(cols_to_delete, reverse=True):
        ws_data.delete_cols(col_idx)

    return wb


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def fill_template(template_wb, rows, globals_data, case_ids, entity_name, country):
    wb = template_wb
    template_type = detect_template_type(wb)
    if template_type == 'MCC':
        wb = fill_mcc(wb, rows, globals_data, case_ids, entity_name, country)
    else:
        wb = fill_cs(wb, rows, globals_data, case_ids, entity_name, country)
    return wb, template_type


def patch_and_save(wb, output_buffer):
    """
    Save workbook to buffer, patching any style alignment corruption
    that can occur after delete_cols operations.
    """
    max_align = len(wb._alignments)
    for xf in wb._cell_styles:
        if xf.alignmentId >= max_align:
            xf.alignmentId = 0
    wb.save(output_buffer)
