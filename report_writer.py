"""
NNS Report Writer - fills MCC and CS templates with processed data.
"""

import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
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
    """Set cell value, skipping merged-cell slaves silently."""
    try:
        ws.cell(row=row, column=col).value = value
    except AttributeError:
        pass

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


# ---------------------------------------------------------------------------
# MCC TEMPLATE FILLER
# ---------------------------------------------------------------------------

def fill_mcc(wb, rows, globals_data, case_ids, entity_name, country):
    ws_summary = wb['LC Summary']
    ws_data    = wb['Data']

    # ---- LC Summary (footer stays at fixed template positions: 20, 23-26, 28) ----
    ws_summary['B8']  = ', '.join(case_ids)
    ws_summary['B9']  = entity_name
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

    # Check for COMPUTER DOMAIN column in summary header row
    SUMMARY_HEADER_ROW = 13
    comp_domain_col_summary = find_col_by_header(ws_summary, SUMMARY_HEADER_ROW, 'COMPUTER DOMAIN')

    # ---- Data sheet ----
    DATA_HEADER_ROW  = 13
    DATA_START_ROW   = 14
    TEMPLATE_DATA_ROWS = 18   # template has 18 pre-bordered rows (14-31)

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

    n_rows = len(rows)

    # Write data into the first n_rows pre-bordered slots
    for idx, row in enumerate(rows):
        r = DATA_START_ROW + idx
        for header, field in mcc_col_order:
            col_idx = col_map.get(header)
            if col_idx is None:
                continue
            val = row.get(field, '-')
            if val is None:
                val = '-'
            # Write dates as datetime objects to preserve Excel date formatting
            if field in ('first_event', 'last_event') and val != '-':
                try:
                    from datetime import datetime, date as date_type
                    if isinstance(val, date_type) and not isinstance(val, datetime):
                        val = datetime(val.year, val.month, val.day)
                except Exception:
                    val = format_date(val)
            safe_set(ws_data, r, col_idx, val)

    # Delete excess pre-bordered rows so footer shifts up naturally
    excess = TEMPLATE_DATA_ROWS - n_rows
    if excess > 0:
        ws_data.delete_rows(DATA_START_ROW + n_rows, excess)

    data_end_row = DATA_START_ROW + n_rows - 1

    # Fix stale cell-level hyperlink refs after delete_rows
    for row in ws_data.iter_rows():
        for cell in row:
            if cell.hyperlink and cell.hyperlink.ref:
                correct_ref = f'{get_column_letter(cell.column)}{cell.row}'
                if cell.hyperlink.ref != correct_ref:
                    cell.hyperlink.ref = correct_ref

    # ---- Column deletion ----
    # Re-read col_map after potential column deletions from Computer Domains check
    cols_to_delete = []
    for header in ['Computer Domains', 'Client Email Addresses']:
        col_idx = col_map.get(header)
        if col_idx and col_all_dash(ws_data, col_idx, DATA_START_ROW, data_end_row):
            cols_to_delete.append(col_idx)

    if col_map.get('Computer Domains') in cols_to_delete and comp_domain_col_summary:
        ws_summary.delete_cols(comp_domain_col_summary)

    for col_idx in sorted(cols_to_delete, reverse=True):
        ws_data.delete_cols(col_idx)

    return wb


# ---------------------------------------------------------------------------
# CS TEMPLATE FILLER
# ---------------------------------------------------------------------------

def _fix_merged_cells_after_row_deletion(ws, deleted_row):
    """
    openpyxl's delete_rows() shifts cell contents but does NOT update merged
    cell range coordinates. This function manually rebuilds the merged cell
    registry, decrementing any row number >= deleted_row by 1.
    """
    old_ranges = [
        (mc.min_row, mc.min_col, mc.max_row, mc.max_col)
        for mc in ws.merged_cells.ranges
    ]
    ws.merged_cells.ranges.clear()
    for (min_row, min_col, max_row, max_col) in old_ranges:
        # Drop merges that were entirely on the deleted row
        if min_row == deleted_row and max_row == deleted_row:
            continue
        # Keep merges entirely above the deleted row unchanged
        if max_row < deleted_row:
            ws.merge_cells(start_row=min_row, start_column=min_col,
                           end_row=max_row, end_column=max_col)
            continue
        # Shift rows at or below the deleted row up by 1
        if min_row >= deleted_row:
            min_row -= 1
        if max_row >= deleted_row:
            max_row -= 1
        ws.merge_cells(start_row=min_row, start_column=min_col,
                       end_row=max_row, end_column=max_col)


def fill_cs(wb, rows, globals_data, case_ids, entity_name, country):
    summary_name = detect_summary_sheet(wb)
    ws_summary   = wb[summary_name]
    ws_data      = wb['Data']

    # ---- Determine if Computer Domain row should be deleted ----
    # Template row 12 = 'Computer Domain'.
    # When no domain data is present, delete that row entirely so all
    # subsequent rows shift up by 1, matching the expected output exactly.
    all_comp_domains = set()
    for row in rows:
        cd = row.get('computer_domain', '-')
        if cd and cd != '-':
            for d in cd.split(','):
                d = d.strip()
                if d:
                    all_comp_domains.add(d)

    delete_domain_row = len(all_comp_domains) == 0
    COMP_DOMAIN_ROW   = 12

    if delete_domain_row:
        # 1. Delete the row
        ws_summary.delete_rows(COMP_DOMAIN_ROW)
        # 2. Fix merged cell coordinates (openpyxl bug: they don't shift automatically)
        _fix_merged_cells_after_row_deletion(ws_summary, COMP_DOMAIN_ROW)
        # 3. Fix stale cell-level hyperlink refs in summary sheet after row deletion
        for summ_row in ws_summary.iter_rows():
            for cell in summ_row:
                if cell.hyperlink and cell.hyperlink.ref:
                    correct_ref = f'{get_column_letter(cell.column)}{cell.row}'
                    if cell.hyperlink.ref != correct_ref:
                        cell.hyperlink.ref = correct_ref
        # Row mappings after deletion (each was original_row - 1)
        ver_row  = 12   # Version row  (was 13)
        yofu_row = 13   # Years of Use (was 14)
        vals_row = 18   # Machines/Users/Events values (was 19)
        lic_row  = 30   # Licensed copies value row (was 31)
        # 3. Update price formula: it referenced C31 which is now C30
        for r in range(25, 35):
            cell = ws_summary.cell(r, 1)
            if cell.value and 'C31' in str(cell.value):
                cell.value = str(cell.value).replace('C31', 'C30')
                break
    else:
        ws_summary['B12'] = ', '.join(sorted(all_comp_domains))
        ver_row  = 13
        yofu_row = 14
        vals_row = 19
        lic_row  = 31

    # ---- Summary sheet: fill all data fields ----
    ws_summary['B9']  = ', '.join(case_ids)
    ws_summary['B10'] = country
    ws_summary['B11'] = entity_name

    # Version string → B{ver_row} (master of B:E merge)
    # Total Versions  → G{ver_row} (free cell, not merged)
    ws_summary.cell(ver_row, 2).value = globals_data['versions_str']
    ws_summary.cell(ver_row, 7).value = globals_data['total_versions']

    # Years of Use → B{yofu_row}
    # Period       → D{yofu_row} (master of D:G merge)
    ws_summary.cell(yofu_row, 2).value = globals_data['years_of_use']
    ws_summary.cell(yofu_row, 4).value = globals_data['period']

    # Machines / Users / Events numeric values
    # B, D, G are masters of their merged spans in the values row
    safe_set(ws_summary, vals_row, 2, globals_data['total_machines'])
    safe_set(ws_summary, vals_row, 4, globals_data['total_users'])
    safe_set(ws_summary, vals_row, 7, globals_data['total_events'])

    # Licensed copies (feeds the price formula)
    safe_set(ws_summary, lic_row, 3, globals_data['total_licenses'])

    # ---- Data sheet: header rows ----
    safe_set(ws_data, 6, 2,  ', '.join(case_ids))
    safe_set(ws_data, 7, 2,  entity_name)
    safe_set(ws_data, 8, 2,  country)
    # DATE: label at col 10, =TODAY() formula at col 11
    safe_set(ws_data, 7, 10, 'DATE:')
    safe_set(ws_data, 7, 11, '=TODAY()')

    DATA_HEADER_ROW    = 11
    DATA_START_ROW     = 12
    TEMPLATE_DATA_ROWS = 12   # 12 pre-bordered rows (12-23) in the blank template

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

    n_rows = len(rows)

    # Write machine data rows — dates as datetime objects (preserves template formatting)
    for idx, row in enumerate(rows):
        r = DATA_START_ROW + idx
        for header, field in cs_col_order:
            col_idx = col_map.get(header)
            if col_idx is None:
                continue
            val = row.get(field, '-')
            if val is None:
                val = '-'
            if field in ('first_event', 'last_event') and val != '-':
                try:
                    from datetime import datetime, date as date_type
                    if isinstance(val, date_type) and not isinstance(val, datetime):
                        val = datetime(val.year, val.month, val.day)
                except Exception:
                    val = format_date(val)
            safe_set(ws_data, r, col_idx, val)

    # Delete excess pre-bordered rows — footer shifts up naturally
    excess = TEMPLATE_DATA_ROWS - n_rows
    if excess > 0:
        ws_data.delete_rows(DATA_START_ROW + n_rows, excess)

    data_end_row = DATA_START_ROW + n_rows - 1

    # Fix stale cell-level hyperlink refs after delete_rows
    for data_row in ws_data.iter_rows():
        for cell in data_row:
            if cell.hyperlink and cell.hyperlink.ref:
                correct_ref = f'{get_column_letter(cell.column)}{cell.row}'
                if cell.hyperlink.ref != correct_ref:
                    cell.hyperlink.ref = correct_ref

    # ---- Column deletion ----
    cols_to_delete = []
    for header in ['Computer Domains', 'Client Email Addresses']:
        col_idx = col_map.get(header)
        if col_idx and col_all_dash(ws_data, col_idx, DATA_START_ROW, data_end_row):
            cols_to_delete.append(col_idx)

    for col_idx in sorted(cols_to_delete, reverse=True):
        ws_data.delete_cols(col_idx)

    return wb


# ---------------------------------------------------------------------------
# MAIN ENTRY POINT
# ---------------------------------------------------------------------------

def fill_template(template_wb, rows, globals_data, case_ids, entity_name, country):
    wb            = template_wb
    template_type = detect_template_type(wb)
    if template_type == 'MCC':
        wb = fill_mcc(wb, rows, globals_data, case_ids, entity_name, country)
    else:
        wb = fill_cs(wb, rows, globals_data, case_ids, entity_name, country)
    return wb, template_type


def patch_and_save(wb, output_buffer):
    """Save workbook, patching style alignment corruption from delete_cols."""
    max_align = len(wb._alignments)
    for xf in wb._cell_styles:
        if xf.alignmentId >= max_align:
            xf.alignmentId = 0
    wb.save(output_buffer)
