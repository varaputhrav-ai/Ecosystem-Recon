"""
Ecosystem Purchase Recon Dashboard
Swiggy Instamart Finance — Month-End Reconciliation
"""

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from datetime import datetime
import re

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────

PACKAGING_CODES = [
    'SFP','SP9','SP3','SPM','STP','SK2','SAP','SP5','SPE','SLP',
    'SHP','SPN','SPP','SIP','SBP','SV1','SKP','PMP','PMM','SPC',
    'PIM','PMC','LKR','PMJ','PMK','SBQ'
]

AMOUNT_TOLERANCE = 1.0   # ₹1 tolerance for amount comparison

# Column identifiers to auto-detect sheet type
SALES_SIGNATURE  = {'INVOICE_NO', 'SELLER_GSTIN', 'CUSTOMER_GSTIN', 'INVOICE_VALUE'}
WMS_SIGNATURE    = {'Invoice_No', 'Supplier_GSTIN', 'Entity_GSTIN', 'Sum of Total_Amt_with_Tax'}
ZOHO_SIGNATURE   = {'Bill Number', 'Vendor GSTIN', 'Entity GSTIN', 'Gross Total'}

# Canonical column names per source
SALES = {
    'inv':          'INVOICE_NO',
    'seller_gstin': 'SELLER_GSTIN',
    'buyer_gstin':  'CUSTOMER_GSTIN',
    'seller_name':  'SELLER_ENTITY',
    'buyer_name':   'BUYER_ENTITY',
    'taxable':      'TAXABLE_VALUE',
    'tax':          'TAX_VALUE',
    'total':        'INVOICE_VALUE',
}
WMS = {
    'inv':          'Invoice_No',
    'seller_gstin': 'Supplier_GSTIN',
    'buyer_gstin':  'Entity_GSTIN',
    'seller_name':  'NEW_SUPPLIER_NAME',
    'buyer_name':   'Inbound Entity',
    'taxable':      'Sum of Total_Amt_without_Tax',
    'tax':          'Sum of Total Tax',
    'total':        'Sum of Total_Amt_with_Tax',
}
ZOHO = {
    'inv':          'Bill Number',
    'seller_gstin': 'Vendor GSTIN',
    'buyer_gstin':  'Entity GSTIN',
    'seller_name':  'Vendor Name',
    'buyer_name':   'Entity',
    'taxable':      '_taxable',      # computed
    'tax':          '_tax',          # computed
    'total':        'Gross Total',
}

COLORS = {
    'hdr_dark':    '1F3864',
    'hdr_mid':     '2F5597',
    'green_light': 'E2EFDA',
    'yellow':      'FFE699',
    'orange':      'FCE4D6',
    'blue_light':  'DDEBF7',
    'balanced':    'C6EFCE',
    'unbalanced':  'FFC7CE',
    'brs_row':     'D6E4F0',
}

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────

def detect_header_row(ws_or_df_rows):
    """Return 0-based index of header row (skip blank leading rows)."""
    for i, row in enumerate(ws_or_df_rows):
        non_empty = [c for c in row if c is not None and str(c).strip() != '']
        if len(non_empty) >= 4:
            return i
    return 0

def load_df(file_obj, sheet_name):
    """Load a sheet, auto-detecting header row."""
    # First pass to find header row
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
    wb.close()
    file_obj.seek(0)

    hdr_idx = detect_header_row(rows)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=hdr_idx, engine='openpyxl')
    df = df.dropna(how='all').dropna(axis=1, how='all')
    # Drop unnamed columns
    df = df[[c for c in df.columns if not str(c).startswith('Unnamed') and not str(c).strip() in ('.', '')]]
    file_obj.seek(0)
    return df

def identify_source(df):
    """Return 'Sales', 'WMS', 'Zoho', or None."""
    cols = set(df.columns.tolist())
    if SALES_SIGNATURE.issubset(cols):
        return 'Sales'
    if WMS_SIGNATURE.issubset(cols):
        return 'WMS'
    if ZOHO_SIGNATURE.issubset(cols):
        return 'Zoho'
    return None

def get_all_sheets_from_file(file_obj):
    """Return dict of {sheet_name: df} for all sheets."""
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    names = wb.sheetnames
    wb.close()
    file_obj.seek(0)
    result = {}
    for name in names:
        try:
            df = load_df(file_obj, name)
            result[name] = df
        except Exception:
            pass
    return result

def prepare(df, source):
    """Normalize invoice numbers, add helper columns."""
    df = df.copy()
    col_map = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[source]

    inv_col = col_map['inv']
    df['_inv'] = df[inv_col].astype(str).str.upper().str.strip()
    df['_gstin_pair'] = (
        df[col_map['seller_gstin']].astype(str).str.strip() + '|' +
        df[col_map['buyer_gstin']].astype(str).str.strip()
    )

    if source == 'Zoho':
        igst = pd.to_numeric(df.get('IGST Amount', 0), errors='coerce').fillna(0)
        cgst = pd.to_numeric(df.get('CGST Amount', 0), errors='coerce').fillna(0)
        sgst = pd.to_numeric(df.get('SGST Amount', 0), errors='coerce').fillna(0)
        gross = pd.to_numeric(df.get('Gross Total', 0), errors='coerce').fillna(0)
        df['_tax']     = (igst + cgst + sgst).round(2)
        df['_taxable'] = (gross - df['_tax']).round(2)

    if source == 'WMS':
        df['_is_pkg'] = df[inv_col].apply(is_packaging)

    return df

def to_num(series):
    return pd.to_numeric(series, errors='coerce').fillna(0)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def is_packaging(inv_no):
    inv = str(inv_no).upper()
    return any(code in inv for code in PACKAGING_CODES)

def norm_amount(val):
    try:
        return round(float(val), 2)
    except Exception:
        return 0.0

# ─────────────────────────────────────────────
# RECON CORE
# ─────────────────────────────────────────────

def run_recon(s1, s1_name, s2, s2_name):
    """
    Returns (matched_df, s1_only_df, s2_only_df).
    matched_df has both sides merged + discrepancy columns.
    """
    cm1 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s1_name]
    cm2 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s2_name]

    # --- Exact invoice match ---
    s1_inv = set(s1['_inv'])
    s2_inv = set(s2['_inv'])

    exact  = s1_inv & s2_inv
    only1  = s1_inv - s2_inv
    only2  = s2_inv - s1_inv

    # Matched
    s1_m = s1[s1['_inv'].isin(exact)].copy()
    s2_m = s2[s2['_inv'].isin(exact)].copy()

    # Deduplicate on invoice (keep first occurrence)
    s1_m = s1_m.drop_duplicates(subset='_inv', keep='first')
    s2_m = s2_m.drop_duplicates(subset='_inv', keep='first')

    # Rename key columns BEFORE merge so pandas doesn't need to add suffixes
    # (pandas only adds suffixes to columns present in BOTH dfs — unreliable
    #  when Sales/WMS/Zoho use different column names for the same concept)
    s1_m = s1_m.rename(columns={
        cm1['total']:        '_m_total_s1',
        cm1['seller_gstin']: '_m_sg_s1',
        cm1['buyer_gstin']:  '_m_bg_s1',
        '_gstin_pair':       '_m_gp_s1',
    })
    s2_m = s2_m.rename(columns={
        cm2['total']:        '_m_total_s2',
        cm2['seller_gstin']: '_m_sg_s2',
        cm2['buyer_gstin']:  '_m_bg_s2',
        '_gstin_pair':       '_m_gp_s2',
    })

    merged = s1_m.merge(s2_m, on='_inv', suffixes=('_s1x', '_s2x'))

    # Amount discrepancy check
    tot1 = to_num(merged['_m_total_s1'])
    tot2 = to_num(merged['_m_total_s2'])
    merged['_amt_diff'] = (tot1 - tot2).round(2)
    merged['_discrepancy'] = merged['_amt_diff'].abs() > AMOUNT_TOLERANCE
    merged['_gstin_diff'] = (
        merged['_m_sg_s1'].astype(str).str.strip() != merged['_m_sg_s2'].astype(str).str.strip()
    ) | (
        merged['_m_bg_s1'].astype(str).str.strip() != merged['_m_bg_s2'].astype(str).str.strip()
    )

    # S1 only
    s1_only = s1[s1['_inv'].isin(only1)].copy()

    # S2 only
    s2_only = s2[s2['_inv'].isin(only2)].copy()

    # --- Secondary match: GSTIN pair + amount (catch wrong inv no in GRN) ---
    if len(s1_only) > 0 and len(s2_only) > 0:
        s2_key_map = {}
        for _, r in s2_only.iterrows():
            key = r['_gstin_pair'] + '|' + str(norm_amount(r[cm2['total']]))
            s2_key_map[key] = s2_key_map.get(key, 0) + 1

        def secondary_category(row):
            key = row['_gstin_pair'] + '|' + str(norm_amount(row[cm1['total']]))
            if key in s2_key_map:
                return 'Invoice No. Mismatch in GRN (Human Error) — Secondary Match Found on GSTIN+Amount'
            return None
        sec = s1_only.apply(secondary_category, axis=1)
        s1_only['_sec_match'] = sec
    else:
        s1_only['_sec_match'] = None

    s1_only['_category'] = s1_only.apply(lambda r: categorize_s1(r, s1_name, s2_name, cm1), axis=1)
    s2_only['_category'] = s2_only.apply(lambda r: categorize_s2(r, s1_name, s2_name, cm2), axis=1)

    return merged, s1_only, s2_only

def categorize_s1(row, s1_name, s2_name, cm):
    # Use secondary match if found
    if row.get('_sec_match'):
        return row['_sec_match']

    if s1_name == 'Sales':
        if s2_name == 'WMS':
            return 'Goods in Transit — GRN not done yet'
        if s2_name == 'Zoho':
            return 'Not Booked in Zoho Books — Provision Required'

    elif s1_name == 'WMS':
        inv = str(row.get('_inv', '')).upper()
        is_pkg = any(code in inv for code in PACKAGING_CODES)
        if s2_name == 'Sales':
            if is_pkg:
                return 'Packing Material Purchase (Not a Sale)'
            return 'GRN Done — Sales Pending / Prior Month Transit'
        if s2_name == 'Zoho':
            if is_pkg:
                return 'Packing Material Purchase'
            return 'GRN Done — Not Booked in Zoho Books'

    elif s1_name == 'Zoho':
        if s2_name == 'Sales':
            return 'Prior Month Sales Booked in Current Month'
        if s2_name == 'WMS':
            return 'Booked in Zoho — GRN Pending'

    return 'To Be Investigated'

def categorize_s2(row, s1_name, s2_name, cm):
    if s2_name == 'WMS':
        inv = str(row.get('_inv', '')).upper()
        if any(code in inv for code in PACKAGING_CODES):
            return 'Packing Material Purchase (Not a Sale)'
        return 'GRN Done — Sales Pending / Prior Month Transit'
    elif s2_name == 'Sales':
        return 'Goods in Transit — GRN Not Done Yet'
    elif s2_name == 'Zoho':
        return 'Prior Month Entry Booked in Zoho'
    return 'To Be Investigated'

# ─────────────────────────────────────────────
# BRS BUILDER
# ─────────────────────────────────────────────

def build_brs(s1, s1_name, s2, s2_name, matched, s1_only, s2_only):
    cm1 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s1_name]
    cm2 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s2_name]

    s1_tot   = to_num(s1[cm1['total']]).sum()
    s2_tot   = to_num(s2[cm2['total']]).sum()
    s1o_tot  = to_num(s1_only[cm1['total']]).sum() if len(s1_only) > 0 else 0
    s2o_tot  = to_num(s2_only[cm2['total']]).sum() if len(s2_only) > 0 else 0

    # Matched value diff (S2 matched total - S1 matched total)
    if len(matched) > 0:
        m_s1_tot = to_num(matched['_m_total_s1']).sum()
        m_s2_tot = to_num(matched['_m_total_s2']).sum()
        matched_val_diff = m_s2_tot - m_s1_tot
    else:
        matched_val_diff = 0

    reconciled = s1_tot - s1o_tot + s2o_tot + matched_val_diff
    diff = round(reconciled - s2_tot, 2)

    rows = [
        ('', f'{s1_name} Total (Invoice Value)',               round(s1_tot, 2),    ''),
        ('(-)', f'In {s1_name} but NOT in {s2_name}',         -round(s1o_tot, 2),  f'{len(s1_only)} invoices'),
        ('(+)', f'In {s2_name} but NOT in {s1_name}',          round(s2o_tot, 2),  f'{len(s2_only)} invoices'),
        ('(+/-)', 'Value diff on matched invoices',             round(matched_val_diff, 2), f'{len(matched[matched["_discrepancy"]]) if len(matched) > 0 else 0} invoices with discrepancy'),
        ('=', f'Reconciled {s2_name} Total',                   round(reconciled, 2), ''),
        ('', f'Actual {s2_name} Total',                        round(s2_tot, 2),    ''),
        ('', 'Difference',                                      diff,                '✓ Balanced' if abs(diff) < AMOUNT_TOLERANCE else '⚠ Unreconciled — Investigate'),
    ]
    return pd.DataFrame(rows, columns=['Sign', 'Particulars', 'Amount (₹)', 'Notes'])

# ─────────────────────────────────────────────
# EXCEL OUTPUT
# ─────────────────────────────────────────────

def apply_fill(ws, row, col_start, col_end, hex_color):
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')
    for c in range(col_start, col_end + 1):
        ws.cell(row=row, column=c).fill = fill

def write_header_row(ws, row, labels, hex_color, font_color='FFFFFF'):
    fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type='solid')
    font = Font(color=font_color, bold=True)
    for c, lbl in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=lbl)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

def write_df_to_ws(ws, df, start_row, fill_color=None, number_cols=None):
    number_cols = number_cols or []
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid') if fill_color else None
    for r_idx, (_, row) in enumerate(df.iterrows()):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=start_row + r_idx, column=c_idx, value=val)
            if fill:
                cell.fill = fill
            if c_idx in number_cols and isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
    return start_row + len(df)

def autofit_columns(ws, max_width=45):
    for col in ws.columns:
        max_len = max((len(str(cell.value or '')) for cell in col if cell.value), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, max_width)

def display_cols_for(source, extra=None):
    cm = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[source]
    cols = [cm['inv'], cm['seller_gstin'], cm['buyer_gstin'],
            cm['seller_name'], cm['buyer_name'],
            cm['taxable'], cm['tax'], cm['total']]
    if extra:
        cols += extra
    return cols

def write_recon_sheet(wb, recon_name, s1_name, s2_name, brs_df, matched, s1_only, s2_only):
    ws = wb.create_sheet(title=recon_name[:31])
    cm1 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s1_name]
    cm2 = {'Sales': SALES, 'WMS': WMS, 'Zoho': ZOHO}[s2_name]
    row = 1

    # ── Title ──
    ws.cell(row=row, column=1, value=f'BRS Reconciliation: {recon_name}').font = Font(bold=True, size=13, color=COLORS['hdr_dark'])
    ws.cell(row=row + 1, column=1, value=f'Generated: {datetime.now().strftime("%d-%b-%Y %H:%M")}').font = Font(italic=True, size=9)
    row += 3

    # ── BRS Table ──
    ws.cell(row=row, column=1, value='BANK RECONCILIATION STATEMENT (BRS)').font = Font(bold=True, size=11)
    row += 1
    write_header_row(ws, row, list(brs_df.columns), COLORS['hdr_dark'])
    row += 1
    for _, brs_row in brs_df.iterrows():
        for c, val in enumerate(brs_row, 1):
            cell = ws.cell(row=row, column=c, value=val)
            if c == 3 and isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
        label = str(brs_row['Particulars'])
        if 'Total' in label or 'Reconciled' in label:
            apply_fill(ws, row, 1, 4, COLORS['brs_row'])
            ws.cell(row=row, column=1).font = Font(bold=True)
            ws.cell(row=row, column=3).font = Font(bold=True)
        if label == 'Difference':
            diff_val = brs_row['Amount (₹)']
            clr = COLORS['balanced'] if abs(float(diff_val or 0)) < AMOUNT_TOLERANCE else COLORS['unbalanced']
            apply_fill(ws, row, 1, 4, clr)
            ws.cell(row=row, column=1).font = Font(bold=True)
        row += 1
    row += 2

    # ── S1 Only ──
    if len(s1_only) > 0:
        s1o_total = to_num(s1_only[cm1['total']]).sum()
        ws.cell(row=row, column=1,
                value=f'In {s1_name} Only — {len(s1_only)} invoices  |  Total: ₹{s1o_total:,.2f}').font = Font(bold=True, size=11, color=COLORS['hdr_dark'])
        row += 1
        show = [c for c in display_cols_for(s1_name) + ['_category'] if c in s1_only.columns]
        write_header_row(ws, row, show, COLORS['hdr_mid'])
        row += 1
        num_pos = [i + 1 for i, c in enumerate(show) if c in (cm1['taxable'], cm1['tax'], cm1['total'])]
        row = write_df_to_ws(ws, s1_only[show], row, fill_color=COLORS['orange'], number_cols=num_pos)
        row += 2

    # ── S2 Only ──
    if len(s2_only) > 0:
        s2o_total = to_num(s2_only[cm2['total']]).sum()
        ws.cell(row=row, column=1,
                value=f'In {s2_name} Only — {len(s2_only)} invoices  |  Total: ₹{s2o_total:,.2f}').font = Font(bold=True, size=11, color=COLORS['hdr_dark'])
        row += 1
        show = [c for c in display_cols_for(s2_name) + ['_category'] if c in s2_only.columns]
        write_header_row(ws, row, show, COLORS['hdr_mid'])
        row += 1
        num_pos = [i + 1 for i, c in enumerate(show) if c in (cm2['taxable'], cm2['tax'], cm2['total'])]
        row = write_df_to_ws(ws, s2_only[show], row, fill_color=COLORS['blue_light'], number_cols=num_pos)
        row += 2

    # ── Matched with Discrepancy ──
    if len(matched) > 0:
        disc = matched[matched['_discrepancy'] | matched['_gstin_diff']].copy()
        if len(disc) > 0:
            ws.cell(row=row, column=1,
                    value=f'Matched — Discrepancies — {len(disc)} invoices').font = Font(bold=True, size=11, color='FF0000')
            row += 1
            tot1_col = '_m_total_s1'
            tot2_col = '_m_total_s2'
            inv_col  = '_inv'
            sg1 = '_m_sg_s1'
            sg2 = '_m_sg_s2'
            bg1 = '_m_bg_s1'
            bg2 = '_m_bg_s2'

            show_cols = [c for c in [inv_col, sg1, sg2, bg1, bg2, tot1_col, tot2_col, '_amt_diff', '_gstin_diff'] if c in disc.columns]
            write_header_row(ws, row, show_cols, COLORS['hdr_mid'])
            row += 1
            num_pos = [i + 1 for i, c in enumerate(show_cols) if c in (tot1_col, tot2_col, '_amt_diff')]
            row = write_df_to_ws(ws, disc[show_cols], row, fill_color=COLORS['yellow'], number_cols=num_pos)

    autofit_columns(ws)
    ws.freeze_panes = 'A4'
    return ws

def create_summary_sheet(wb, results):
    ws = wb.create_sheet(title='Summary', index=0)
    ws.cell(row=1, column=1, value='Ecosystem Purchase Recon — Monthly Summary').font = Font(bold=True, size=14, color=COLORS['hdr_dark'])
    ws.cell(row=2, column=1, value=f'Generated: {datetime.now().strftime("%d-%b-%Y %H:%M")}').font = Font(italic=True, size=9)

    headers = ['Recon', 'S1 Invoice Total (₹)', 'S2 Invoice Total (₹)',
               'In S1 Only (Count)', 'In S2 Only (Count)',
               'Value Discrepancies (Count)', 'BRS Difference (₹)', 'Status']
    write_header_row(ws, 4, headers, COLORS['hdr_dark'])

    for i, (name, brs, s1_only, s2_only, matched) in enumerate(results, 5):
        diff      = brs.iloc[6]['Amount (₹)']
        disc_cnt  = len(matched[matched['_discrepancy']]) if len(matched) > 0 else 0
        status    = '✓ Balanced' if abs(float(diff or 0)) < AMOUNT_TOLERANCE else f'⚠ Diff ₹{diff:,.2f}'

        row_data = [
            name,
            round(brs.iloc[0]['Amount (₹)'], 2),
            round(brs.iloc[5]['Amount (₹)'], 2),
            len(s1_only),
            len(s2_only),
            disc_cnt,
            round(float(diff), 2),
            status,
        ]
        for c, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=c, value=val)
            if c in (2, 3, 7):
                cell.number_format = '#,##0.00'
        clr = COLORS['balanced'] if abs(float(diff or 0)) < AMOUNT_TOLERANCE else COLORS['unbalanced']
        apply_fill(ws, i, 8, 8, clr)

    autofit_columns(ws)
    ws.freeze_panes = 'A5'

def generate_excel(results, s1_map):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    create_summary_sheet(wb, results)
    for name, brs, s1_only, s2_only, matched in results:
        s1n, s2n = name.split(' vs ')
        write_recon_sheet(wb, name, s1n, s2n, brs, matched, s1_only, s2_only)
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title='Ecosystem Purchase Recon',
        page_icon='📊',
        layout='wide',
        initial_sidebar_state='collapsed',
    )

    # Custom CSS
    st.markdown("""
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 1rem; }
    .stButton > button { border-radius: 6px; font-weight: 600; }
    .stDownloadButton > button { background: #1F3864; color: white; border-radius: 6px; font-weight: 600; font-size: 1rem; padding: 0.6rem 1.5rem; }
    .metric-card { background: #f0f4ff; border-radius: 8px; padding: 1rem 1.5rem; margin: 0.25rem; }
    </style>
    """, unsafe_allow_html=True)

    st.title('📊 Ecosystem Purchase Recon')
    st.caption('Swiggy Instamart Finance | Month-End Reconciliation Dashboard')
    st.divider()

    # ── Upload Section ──
    st.subheader('Step 1 — Upload Source Files')
    st.markdown('Upload files one by one **or** a single combined workbook that has all 3 sheets.')

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('**📤 Sales Data**')
        sales_file = st.file_uploader('Sales Excel', type=['xlsx', 'xls'], key='sales', label_visibility='collapsed')
        if sales_file:
            st.success(f'✓ {sales_file.name}')
    with col2:
        st.markdown('**📦 WMS / Purchase Data**')
        wms_file = st.file_uploader('WMS Excel', type=['xlsx', 'xls'], key='wms', label_visibility='collapsed')
        if wms_file:
            st.success(f'✓ {wms_file.name}')
    with col3:
        st.markdown('**📒 Zoho Books Data**')
        zoho_file = st.file_uploader('Zoho Excel', type=['xlsx', 'xls'], key='zoho', label_visibility='collapsed')
        if zoho_file:
            st.success(f'✓ {zoho_file.name}')

    st.divider()

    # ── Run ──
    run_disabled = not (sales_file or wms_file or zoho_file)
    run_btn = st.button('▶  Run Reconciliation', type='primary', use_container_width=True, disabled=run_disabled)

    if not run_btn:
        if run_disabled:
            st.info('👆 Upload at least one file to begin. You can also upload a single combined workbook.')
        return

    # ── Load & Identify Sources ──
    dfs = {}   # source_name → dataframe

    with st.spinner('Loading files…'):
        for label, fobj in [('Sales', sales_file), ('WMS', wms_file), ('Zoho', zoho_file)]:
            if fobj is None:
                continue
            sheets = get_all_sheets_from_file(fobj)
            for sname, df in sheets.items():
                src = identify_source(df)
                if src and src not in dfs:
                    dfs[src] = df
                    st.info(f'Detected **{src}** in "{fobj.name}" → sheet "{sname}"')

    # Check we have all 3
    missing = [s for s in ('Sales', 'WMS', 'Zoho') if s not in dfs]
    if missing:
        st.error(f'Could not identify: {", ".join(missing)}. Please check column names match the expected format.')
        with st.expander('Expected column names'):
            c1, c2, c3 = st.columns(3)
            c1.markdown('**Sales:** ' + ', '.join(SALES_SIGNATURE))
            c2.markdown('**WMS:** ' + ', '.join(WMS_SIGNATURE))
            c3.markdown('**Zoho:** ' + ', '.join(ZOHO_SIGNATURE))
        return

    with st.spinner('Preparing data…'):
        sales_df = prepare(dfs['Sales'], 'Sales')
        wms_df   = prepare(dfs['WMS'],   'WMS')
        zoho_df  = prepare(dfs['Zoho'],  'Zoho')

    st.success(f'Loaded — Sales: {len(sales_df):,} rows | WMS: {len(wms_df):,} rows | Zoho: {len(zoho_df):,} rows')

    # ── Run 6 Recons ──
    st.subheader('Step 2 — Running Reconciliations')
    progress_bar = st.progress(0)
    status_text  = st.empty()

    PAIRS = [
        ('Sales', 'WMS',   sales_df, wms_df),
        ('WMS',   'Sales', wms_df,   sales_df),
        ('Sales', 'Zoho',  sales_df, zoho_df),
        ('Zoho',  'Sales', zoho_df,  sales_df),
        ('WMS',   'Zoho',  wms_df,   zoho_df),
        ('Zoho',  'WMS',   zoho_df,  wms_df),
    ]

    results = []
    for i, (s1n, s2n, s1, s2) in enumerate(PAIRS):
        status_text.text(f'Running {s1n} vs {s2n}…')
        matched, s1_only, s2_only = run_recon(s1, s1n, s2, s2n)
        brs = build_brs(s1, s1n, s2, s2n, matched, s1_only, s2_only)
        results.append((f'{s1n} vs {s2n}', brs, s1_only, s2_only, matched))
        progress_bar.progress((i + 1) / len(PAIRS))

    status_text.empty()

    # ── Summary Table ──
    st.subheader('Step 3 — Summary')
    summary_rows = []
    for name, brs, s1_only, s2_only, matched in results:
        diff     = brs.iloc[6]['Amount (₹)']
        disc_cnt = int(matched['_discrepancy'].sum()) if len(matched) > 0 else 0
        summary_rows.append({
            'Recon':              name,
            'S1 Total (₹)':      f"{brs.iloc[0]['Amount (₹)']:,.2f}",
            'S2 Total (₹)':      f"{brs.iloc[5]['Amount (₹)']:,.2f}",
            'S1 Only (invoices)': len(s1_only),
            'S2 Only (invoices)': len(s2_only),
            'Value Discrepancies':disc_cnt,
            'Status':            '✅ Balanced' if abs(float(diff or 0)) < AMOUNT_TOLERANCE else f'⚠️ Diff ₹{diff:,.2f}',
        })

    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

    # Metrics row
    total_s1_only  = sum(len(r[2]) for r in results[:3:2])   # Sales vs WMS, Sales vs Zoho
    total_s2_only  = sum(len(r[3]) for r in results[:3:2])
    total_disc     = sum(int(r[4]['_discrepancy'].sum()) if len(r[4]) > 0 else 0 for r in results)
    balanced_count = sum(1 for r in results if abs(float(r[1].iloc[6]['Amount (₹)'] or 0)) < AMOUNT_TOLERANCE)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric('Recons Balanced', f'{balanced_count} / 6')
    m2.metric('S1-Only Invoices (unique)', total_s1_only)
    m3.metric('S2-Only Invoices (unique)', total_s2_only)
    m4.metric('Value Discrepancies', total_disc)

    # ── Export ──
    st.divider()
    st.subheader('Step 4 — Download Output')

    with st.spinner('Building Excel workbook…'):
        excel_buf = generate_excel(results, None)

    month_str = datetime.now().strftime('%b_%Y')
    st.download_button(
        label='📥 Download Recon Output (Excel)',
        data=excel_buf,
        file_name=f'Ecosystem_Recon_{month_str}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )

    st.caption('Output has 7 sheets: Summary + one BRS sheet per recon. Each BRS sheet shows the reconciliation statement, S1-only items with categories, S2-only items with categories, and matched invoices with discrepancies.')

if __name__ == '__main__':
    main()
