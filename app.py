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

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────

PACKAGING_CODES = [
    'SFP','SP9','SP3','SPM','STP','SK2','SAP','SP5','SPE','SLP',
    'SHP','SPN','SPP','SIP','SBP','SV1','SKP','PMP','PMM','SPC',
    'PIM','PMC','LKR','PMJ','PMK','SBQ'
]

AMOUNT_TOLERANCE = 1.0

SALES_CANDIDATES = {
    'inv':          ['INVOICE_NO'],
    'seller_gstin': ['SELLER_GSTIN'],
    'buyer_gstin':  ['CUSTOMER_GSTIN'],
    'seller_name':  ['SELLER_ENTITY'],
    'buyer_name':   ['BUYER_ENTITY', 'ENTITY'],
    'taxable':      ['TAXABLE_VALUE'],
    'tax':          ['TAX_VALUE'],
    'total':        ['Amount for Reco', 'INVOICE_VALUE'],
}
WMS_CANDIDATES = {
    'inv':          ['Invoice_No'],
    'seller_gstin': ['Supplier_GSTIN', 'Supplier_GSTN'],
    'buyer_gstin':  ['Entity_GSTIN', 'Entity_GSTN'],
    'seller_name':  ['NEW_SUPPLIER_NAME', 'Vendor Name'],
    'buyer_name':   ['Inbound Entity', 'Entity'],
    'taxable':      ['Sum of Total_Amt_without_Tax'],
    'tax':          ['Sum of Total Tax', 'ITEM_TAX_VALUE'],
    'total':        ['Sum of Total_Amt_with_Tax'],
}
ZOHO_CANDIDATES = {
    'inv':          ['Bill Number'],
    'seller_gstin': ['Vendor GSTIN'],
    'buyer_gstin':  ['Entity GSTIN'],
    'seller_name':  ['Vendor Name'],
    'buyer_name':   ['Entity'],
    'taxable':      ['_taxable'],
    'tax':          ['_tax'],
    'total':        ['Gross Total', 'Invoice Total', 'Sum of Item Total'],
}

SALES_SIG = [['INVOICE_NO'], ['SELLER_GSTIN'], ['CUSTOMER_GSTIN'], ['INVOICE_VALUE', 'Amount for Reco']]
WMS_SIG   = [['Invoice_No'], ['Supplier_GSTIN', 'Supplier_GSTN'], ['Entity_GSTIN', 'Entity_GSTN'], ['Sum of Total_Amt_with_Tax']]
ZOHO_SIG  = [['Bill Number'], ['Vendor GSTIN'], ['Entity GSTIN'], ['Gross Total', 'Invoice Total', 'Sum of Item Total']]

COLORS = {
    'hdr_dark':  '1F3864', 'hdr_mid': '2F5597',
    'orange':    'FCE4D6', 'blue_light': 'DDEBF7',
    'yellow':    'FFE699', 'balanced': 'C6EFCE',
    'unbalanced':'FFC7CE', 'brs_row':  'D6E4F0',
}

# ─────────────────────────────────────────────
# COLUMN RESOLUTION
# ─────────────────────────────────────────────

def find_col(df, *candidates):
    cols_stripped = {str(c).strip(): str(c) for c in df.columns}
    for c in candidates:
        if str(c).strip() in cols_stripped:
            return cols_stripped[str(c).strip()]
    return None

def resolve_cols(df, candidates_dict):
    result = {}
    for key, options in candidates_dict.items():
        if options[0].startswith('_'):
            result[key] = options[0]
            continue
        result[key] = find_col(df, *options)
    return result

def matches_sig(df, sig_groups):
    cols = {str(c).strip() for c in df.columns}
    for group in sig_groups:
        if not any(str(c).strip() in cols for c in group):
            return False
    return True

def identify_source(df):
    if matches_sig(df, SALES_SIG): return 'Sales'
    if matches_sig(df, WMS_SIG):   return 'WMS'
    if matches_sig(df, ZOHO_SIG):  return 'Zoho'
    return None

# ─────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────

def detect_header_row(rows):
    for i, row in enumerate(rows):
        if len([c for c in row if c is not None and str(c).strip()]) >= 4:
            return i
    return 0

def load_df(file_obj, sheet_name):
    wb = openpyxl.load_workbook(file_obj, read_only=True, data_only=True)
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(min_row=1, max_row=10, values_only=True))
    wb.close()
    file_obj.seek(0)
    hdr_idx = detect_header_row(rows)
    df = pd.read_excel(file_obj, sheet_name=sheet_name, header=hdr_idx, engine='openpyxl')
    df = df.dropna(how='all').dropna(axis=1, how='all')
    df.columns = [str(c).strip() for c in df.columns]
    df = df[[c for c in df.columns if not c.startswith('Unnamed') and c not in ('.', '', 'None')]]
    file_obj.seek(0)
    return df

def get_all_sheets(file_obj):
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

def to_num(s):
    return pd.to_numeric(s, errors='coerce').fillna(0)

def is_packaging(inv_no):
    inv = str(inv_no).upper()
    return any(code in inv for code in PACKAGING_CODES)

def norm_amount(val):
    try: return round(float(val), 2)
    except: return 0.0

def prepare(df, source):
    df = df.copy()
    cands = {'Sales': SALES_CANDIDATES, 'WMS': WMS_CANDIDATES, 'Zoho': ZOHO_CANDIDATES}[source]
    cm = resolve_cols(df, cands)
    df.attrs['_cm'] = cm
    df.attrs['_source'] = source

    df['_inv'] = df[cm['inv']].astype(str).str.upper().str.strip()
    df['_gstin_pair'] = (
        df[cm['seller_gstin']].astype(str).str.strip() + '|' +
        df[cm['buyer_gstin']].astype(str).str.strip()
    )

    if source == 'Zoho':
        total_col = cm['total']
        gross = to_num(df[total_col]) if total_col else pd.Series(0, index=df.index)
        igst_col = find_col(df, 'IGST Amount', 'IGST')
        cgst_col = find_col(df, 'CGST Amount', 'CGST')
        sgst_col = find_col(df, 'SGST Amount', 'SGST')
        tax_col  = find_col(df, 'Sum of Tax Amount', 'Tax Amount', 'Total Tax')
        igst = to_num(df[igst_col]) if igst_col else pd.Series(0, index=df.index)
        cgst = to_num(df[cgst_col]) if cgst_col else pd.Series(0, index=df.index)
        sgst = to_num(df[sgst_col]) if sgst_col else pd.Series(0, index=df.index)
        if (igst + cgst + sgst).sum() > 0:
            df['_tax'] = (igst + cgst + sgst).round(2)
        elif tax_col:
            df['_tax'] = to_num(df[tax_col]).round(2)
        else:
            df['_tax'] = 0.0
        df['_taxable'] = (gross - df['_tax']).round(2)

    if source == 'WMS':
        df['_is_pkg'] = df[cm['inv']].apply(is_packaging)

    return df

def get_cm(df):
    return df.attrs.get('_cm', {})

# ─────────────────────────────────────────────
# RECON CORE
# ─────────────────────────────────────────────

def run_recon(s1, s1_name, s2, s2_name):
    cm1 = get_cm(s1)
    cm2 = get_cm(s2)

    s1_inv = set(s1['_inv'])
    s2_inv = set(s2['_inv'])
    exact = s1_inv & s2_inv
    only1 = s1_inv - s2_inv
    only2 = s2_inv - s1_inv

    s1_m = s1[s1['_inv'].isin(exact)].drop_duplicates('_inv').copy()
    s2_m = s2[s2['_inv'].isin(exact)].drop_duplicates('_inv').copy()

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
    merged['_amt_diff']   = (to_num(merged['_m_total_s1']) - to_num(merged['_m_total_s2'])).round(2)
    merged['_discrepancy'] = merged['_amt_diff'].abs() > AMOUNT_TOLERANCE
    merged['_gstin_diff']  = (
        merged['_m_sg_s1'].astype(str).str.strip() != merged['_m_sg_s2'].astype(str).str.strip()
    ) | (
        merged['_m_bg_s1'].astype(str).str.strip() != merged['_m_bg_s2'].astype(str).str.strip()
    )

    s1_only = s1[s1['_inv'].isin(only1)].copy()
    s2_only = s2[s2['_inv'].isin(only2)].copy()

    # Secondary match: GSTIN pair + amount (catch wrong invoice no in GRN)
    if len(s1_only) > 0 and len(s2_only) > 0:
        s2_keys = set(
            r['_gstin_pair'] + '|' + str(norm_amount(r[cm2['total']]))
            for _, r in s2_only.iterrows()
        )
        def sec_cat(row):
            key = row['_gstin_pair'] + '|' + str(norm_amount(row[cm1['total']]))
            if key in s2_keys:
                return 'Invoice No. Mismatch in GRN (Human Error) — Secondary Match on GSTIN+Amount'
            return None
        s1_only['_sec_match'] = s1_only.apply(sec_cat, axis=1)
    else:
        s1_only['_sec_match'] = None

    s1_only['_category'] = s1_only.apply(lambda r: categorize_s1(r, s1_name, s2_name), axis=1)
    s2_only['_category'] = s2_only.apply(lambda r: categorize_s2(r, s1_name, s2_name), axis=1)

    return merged, s1_only, s2_only

def categorize_s1(row, s1_name, s2_name):
    if row.get('_sec_match'):
        return row['_sec_match']
    if s1_name == 'Sales':
        if s2_name == 'WMS':  return 'Goods in Transit — GRN Not Done Yet'
        if s2_name == 'Zoho': return 'Not Booked in Zoho Books — Provision Required'
    elif s1_name == 'WMS':
        if any(code in str(row.get('_inv','')).upper() for code in PACKAGING_CODES):
            return 'Packing Material Purchase (Not a Sale)'
        if s2_name == 'Sales': return 'GRN Done — Sales Pending / Prior Month Transit'
        if s2_name == 'Zoho':  return 'GRN Done — Not Booked in Zoho Books'
    elif s1_name == 'Zoho':
        if s2_name == 'Sales': return 'Prior Month Sales Booked in Current Month'
        if s2_name == 'WMS':   return 'Booked in Zoho — GRN Pending'
    return 'To Be Investigated'

def categorize_s2(row, s1_name, s2_name):
    if s2_name == 'WMS':
        if any(code in str(row.get('_inv','')).upper() for code in PACKAGING_CODES):
            return 'Packing Material Purchase (Not a Sale)'
        return 'GRN Done — Sales Pending / Prior Month Transit'
    elif s2_name == 'Sales': return 'Goods in Transit — GRN Not Done Yet'
    elif s2_name == 'Zoho':  return 'Prior Month Entry Booked in Zoho'
    return 'To Be Investigated'

def build_brs(s1, s1_name, s2, s2_name, matched, s1_only, s2_only):
    cm1 = get_cm(s1)
    cm2 = get_cm(s2)
    s1_tot  = to_num(s1[cm1['total']]).sum()
    s2_tot  = to_num(s2[cm2['total']]).sum()
    s1o_tot = to_num(s1_only[cm1['total']]).sum() if len(s1_only) > 0 else 0
    s2o_tot = to_num(s2_only[cm2['total']]).sum() if len(s2_only) > 0 else 0
    m_diff  = (to_num(matched['_m_total_s2']).sum() - to_num(matched['_m_total_s1']).sum()) if len(matched) > 0 else 0
    reconciled = s1_tot - s1o_tot + s2o_tot + m_diff
    diff = round(reconciled - s2_tot, 2)
    disc_cnt = int(matched['_discrepancy'].sum()) if len(matched) > 0 else 0
    rows = [
        ('',     f'{s1_name} Total',                   round(s1_tot, 2),    ''),
        ('(-)',  f'In {s1_name} only (not in {s2_name})', -round(s1o_tot,2), f'{len(s1_only)} invoices'),
        ('(+)',  f'In {s2_name} only (not in {s1_name})',  round(s2o_tot,2), f'{len(s2_only)} invoices'),
        ('(±)',  'Value diff on matched invoices',      round(m_diff, 2),    f'{disc_cnt} invoices with discrepancy'),
        ('=',   f'Reconciled {s2_name} Total',         round(reconciled,2), ''),
        ('',    f'Actual {s2_name} Total',             round(s2_tot, 2),    ''),
        ('',    'Difference',                           diff,                '✓ Balanced' if abs(diff) < AMOUNT_TOLERANCE else '⚠ Investigate'),
    ]
    return pd.DataFrame(rows, columns=['', 'Particulars', 'Amount (₹)', 'Notes'])

# ─────────────────────────────────────────────
# DISPLAY HELPERS
# ─────────────────────────────────────────────

def display_cols(cm, df):
    """Return display-worthy columns that exist in df."""
    keys = ['inv','seller_name','buyer_name','seller_gstin','buyer_gstin','taxable','tax','total']
    cols = []
    for k in keys:
        v = cm.get(k)
        if v and v in df.columns:
            cols.append(v)
    return cols

def style_brs(brs_df):
    def row_style(row):
        label = str(row['Particulars'])
        diff  = row['Amount (₹)']
        if label == 'Difference':
            color = '#C6EFCE' if abs(float(diff or 0)) < AMOUNT_TOLERANCE else '#FFC7CE'
            return [f'background-color:{color};font-weight:bold'] * len(row)
        if 'Total' in label or 'Reconciled' in label or label.startswith('='):
            return ['background-color:#D6E4F0;font-weight:bold'] * len(row)
        return [''] * len(row)
    return brs_df.style.apply(row_style, axis=1).format({'Amount (₹)': '{:,.2f}'})

def render_table(df, cm, height=400, key=None):
    """Render a dataframe with only meaningful columns."""
    show = display_cols(cm, df)
    cat_col = '_category' if '_category' in df.columns else None
    if cat_col:
        show = show + [cat_col]
    show = [c for c in show if c in df.columns]
    if not show:
        st.info('No columns to display.')
        return
    rename_map = {
        cm.get('inv'):          'Invoice No.',
        cm.get('seller_name'):  'Seller',
        cm.get('buyer_name'):   'Buyer',
        cm.get('seller_gstin'): 'Seller GSTIN',
        cm.get('buyer_gstin'):  'Buyer GSTIN',
        cm.get('taxable'):      'Taxable Value (₹)',
        cm.get('tax'):          'Tax Value (₹)',
        cm.get('total'):        'Invoice Total (₹)',
        '_category':            'Category / Reason',
    }
    display_df = df[show].rename(columns={k: v for k, v in rename_map.items() if k})
    st.dataframe(display_df, use_container_width=True, height=height, hide_index=True)

# ─────────────────────────────────────────────
# EXCEL EXPORT
# ─────────────────────────────────────────────

def apply_fill(ws, row, c1, c2, color):
    fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
    for c in range(c1, c2 + 1):
        ws.cell(row=row, column=c).fill = fill

def write_hdr(ws, row, labels, color, font_color='FFFFFF'):
    fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
    for c, lbl in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=lbl)
        cell.fill = fill
        cell.font = Font(color=font_color, bold=True)
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

def write_rows(ws, df, start, fill_color=None, num_cols=None):
    num_cols = num_cols or []
    fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid') if fill_color else None
    for ri, (_, row) in enumerate(df.iterrows()):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=start + ri, column=ci, value=val)
            if fill: cell.fill = fill
            if ci in num_cols and isinstance(val, (int, float)):
                cell.number_format = '#,##0.00'
    return start + len(df)

def autofit(ws):
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col if c.value), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(w + 3, 45)

def build_excel(results):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Summary sheet
    ws = wb.create_sheet('Summary', 0)
    ws.cell(1,1,'Ecosystem Purchase Recon — Monthly Summary').font = Font(bold=True, size=14, color=COLORS['hdr_dark'])
    ws.cell(2,1, f'Generated: {datetime.now().strftime("%d-%b-%Y %H:%M")}').font = Font(italic=True, size=9)
    hdrs = ['Recon','S1 Total (₹)','S2 Total (₹)','S1 Only (count)','S2 Only (count)','Discrepancies','Difference (₹)','Status']
    write_hdr(ws, 4, hdrs, COLORS['hdr_dark'])
    for i, (name, brs, s1_only, s2_only, matched, cm1, cm2) in enumerate(results, 5):
        diff = brs.iloc[6]['Amount (₹)']
        disc = int(matched['_discrepancy'].sum()) if len(matched) > 0 else 0
        vals = [name, brs.iloc[0]['Amount (₹)'], brs.iloc[5]['Amount (₹)'],
                len(s1_only), len(s2_only), disc, float(diff),
                '✓ Balanced' if abs(float(diff or 0)) < AMOUNT_TOLERANCE else f'⚠ Diff ₹{diff:,.2f}']
        for c, v in enumerate(vals, 1):
            cell = ws.cell(i, c, v)
            if c in (2,3,7): cell.number_format = '#,##0.00'
        clr = COLORS['balanced'] if abs(float(diff or 0)) < AMOUNT_TOLERANCE else COLORS['unbalanced']
        apply_fill(ws, i, 8, 8, clr)
    autofit(ws)
    ws.freeze_panes = 'A5'

    # Recon sheets
    for name, brs, s1_only, s2_only, matched, cm1, cm2 in results:
        s1n, s2n = name.split(' vs ')
        ws = wb.create_sheet(name[:31])
        r = 1

        ws.cell(r,1, f'BRS: {name}').font = Font(bold=True, size=13, color=COLORS['hdr_dark'])
        ws.cell(r+1, 1, f'Generated: {datetime.now().strftime("%d-%b-%Y %H:%M")}').font = Font(italic=True, size=9)
        r += 3

        # BRS table
        ws.cell(r,1,'RECONCILIATION STATEMENT').font = Font(bold=True, size=11)
        r += 1
        write_hdr(ws, r, list(brs.columns), COLORS['hdr_dark'])
        r += 1
        for _, brow in brs.iterrows():
            for c, val in enumerate(brow, 1):
                cell = ws.cell(r, c, val)
                if c == 3 and isinstance(val, (int,float)): cell.number_format = '#,##0.00'
            lbl = str(brow['Particulars'])
            if 'Total' in lbl or 'Reconciled' in lbl:
                apply_fill(ws, r, 1, 4, COLORS['brs_row'])
                ws.cell(r,1).font = Font(bold=True)
                ws.cell(r,3).font = Font(bold=True)
            if lbl == 'Difference':
                diff_v = brow['Amount (₹)']
                apply_fill(ws, r, 1, 4, COLORS['balanced'] if abs(float(diff_v or 0)) < AMOUNT_TOLERANCE else COLORS['unbalanced'])
                ws.cell(r,1).font = Font(bold=True)
            r += 1
        r += 2

        # S1 only
        if len(s1_only) > 0:
            tot = to_num(s1_only[cm1['total']]).sum() if cm1.get('total') and cm1['total'] in s1_only.columns else 0
            ws.cell(r,1, f'In {s1n} Only — {len(s1_only)} invoices | Total: ₹{tot:,.2f}').font = Font(bold=True, size=11, color=COLORS['hdr_dark'])
            r += 1
            show = display_cols(cm1, s1_only) + (['_category'] if '_category' in s1_only.columns else [])
            write_hdr(ws, r, show, COLORS['hdr_mid'])
            r += 1
            num_p = [i+1 for i,c in enumerate(show) if c in (cm1.get('taxable'), cm1.get('tax'), cm1.get('total'))]
            r = write_rows(ws, s1_only[show], r, COLORS['orange'], num_p)
            r += 2

        # S2 only
        if len(s2_only) > 0:
            tot = to_num(s2_only[cm2['total']]).sum() if cm2.get('total') and cm2['total'] in s2_only.columns else 0
            ws.cell(r,1, f'In {s2n} Only — {len(s2_only)} invoices | Total: ₹{tot:,.2f}').font = Font(bold=True, size=11, color=COLORS['hdr_dark'])
            r += 1
            show = display_cols(cm2, s2_only) + (['_category'] if '_category' in s2_only.columns else [])
            write_hdr(ws, r, show, COLORS['hdr_mid'])
            r += 1
            num_p = [i+1 for i,c in enumerate(show) if c in (cm2.get('taxable'), cm2.get('tax'), cm2.get('total'))]
            r = write_rows(ws, s2_only[show], r, COLORS['blue_light'], num_p)
            r += 2

        # Discrepancies
        if len(matched) > 0:
            disc = matched[matched['_discrepancy'] | matched['_gstin_diff']].copy()
            if len(disc) > 0:
                ws.cell(r,1, f'Matched — Discrepancies — {len(disc)} invoices').font = Font(bold=True, size=11, color='FF0000')
                r += 1
                disc_cols = [c for c in ['_inv','_m_sg_s1','_m_sg_s2','_m_bg_s1','_m_bg_s2','_m_total_s1','_m_total_s2','_amt_diff','_gstin_diff'] if c in disc.columns]
                write_hdr(ws, r, disc_cols, COLORS['hdr_mid'])
                r += 1
                num_p = [i+1 for i,c in enumerate(disc_cols) if 'total' in c or 'diff' in c]
                r = write_rows(ws, disc[disc_cols], r, COLORS['yellow'], num_p)

        autofit(ws)
        ws.freeze_panes = 'A4'

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

    st.markdown("""
    <style>
    .block-container{padding-top:1.2rem;padding-bottom:1rem}
    .stDownloadButton>button{background:#1F3864;color:white;border-radius:6px;font-weight:600}
    .stTabs [data-baseweb="tab"]{font-weight:600}
    div[data-testid="metric-container"]>div{font-size:0.85rem}
    </style>""", unsafe_allow_html=True)

    st.title('📊 Ecosystem Purchase Recon')
    st.caption('Swiggy Instamart Finance | Month-End Reconciliation')
    st.divider()

    # ── Upload ──────────────────────────────────
    st.subheader('Step 1 — Upload Source Files')
    st.markdown('Upload files separately **or** a single combined workbook with all 3 sheets.')
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown('**📤 Sales Data**')
        sales_file = st.file_uploader('Sales', type=['xlsx','xls'], key='sales', label_visibility='collapsed')
        if sales_file: st.success(f'✓ {sales_file.name}')
    with c2:
        st.markdown('**📦 WMS / Purchase Data**')
        wms_file = st.file_uploader('WMS', type=['xlsx','xls'], key='wms', label_visibility='collapsed')
        if wms_file: st.success(f'✓ {wms_file.name}')
    with c3:
        st.markdown('**📒 Zoho Books Data**')
        zoho_file = st.file_uploader('Zoho', type=['xlsx','xls'], key='zoho', label_visibility='collapsed')
        if zoho_file: st.success(f'✓ {zoho_file.name}')

    st.divider()
    if not st.button('▶  Run Reconciliation', type='primary', use_container_width=True,
                     disabled=not any([sales_file, wms_file, zoho_file])):
        st.info('👆 Upload at least one file to begin.')
        return

    # ── Load & Detect ────────────────────────────
    dfs = {}
    with st.spinner('Loading files…'):
        for _label, fobj in [('Sales', sales_file), ('WMS', wms_file), ('Zoho', zoho_file)]:
            if fobj is None: continue
            for sname, df in get_all_sheets(fobj).items():
                src = identify_source(df)
                if src and src not in dfs:
                    dfs[src] = df
                    st.info(f'Detected **{src}** in "{fobj.name}" → sheet "{sname}"')

    missing = [s for s in ('Sales','WMS','Zoho') if s not in dfs]
    if missing:
        st.error(f'Could not identify: {", ".join(missing)}. Check column names.')
        with st.expander('Expected column names'):
            st.markdown('**Sales:** INVOICE_NO, SELLER_GSTIN, CUSTOMER_GSTIN, INVOICE_VALUE / Amount for Reco')
            st.markdown('**WMS:** Invoice_No, Supplier_GSTIN/GSTN, Entity_GSTIN/GSTN, Sum of Total_Amt_with_Tax')
            st.markdown('**Zoho:** Bill Number, Vendor GSTIN, Entity GSTIN, Gross Total / Invoice Total')
        return

    with st.spinner('Preparing data…'):
        sales_df = prepare(dfs['Sales'], 'Sales')
        wms_df   = prepare(dfs['WMS'],   'WMS')
        zoho_df  = prepare(dfs['Zoho'],  'Zoho')

    st.success(f'Loaded — Sales: {len(sales_df):,} | WMS: {len(wms_df):,} | Zoho: {len(zoho_df):,} rows')

    # ── Run Recons ──────────────────────────────
    PAIRS = [
        ('Sales','WMS',  sales_df, wms_df),
        ('WMS','Sales',  wms_df,   sales_df),
        ('Sales','Zoho', sales_df, zoho_df),
        ('Zoho','Sales', zoho_df,  sales_df),
        ('WMS','Zoho',   wms_df,   zoho_df),
        ('Zoho','WMS',   zoho_df,  wms_df),
    ]
    prog  = st.progress(0)
    stxt  = st.empty()
    results = []
    for i, (s1n, s2n, s1, s2) in enumerate(PAIRS):
        stxt.text(f'Running {s1n} vs {s2n}…')
        matched, s1_only, s2_only = run_recon(s1, s1n, s2, s2n)
        brs = build_brs(s1, s1n, s2, s2n, matched, s1_only, s2_only)
        results.append((f'{s1n} vs {s2n}', brs, s1_only, s2_only, matched, get_cm(s1), get_cm(s2)))
        prog.progress((i+1)/len(PAIRS))
    stxt.empty()

    # ── Summary Banner ──────────────────────────
    st.divider()
    st.subheader('📋 Recon Summary')
    mcols = st.columns(6)
    for col, (name, brs, s1_only, s2_only, matched, cm1, cm2) in zip(mcols, results):
        diff = float(brs.iloc[6]['Amount (₹)'] or 0)
        balanced = abs(diff) < AMOUNT_TOLERANCE
        col.metric(
            label=name,
            value='✅ Balanced' if balanced else '⚠️ Diff',
            delta=f'₹{diff:,.0f}' if not balanced else f'{len(s1_only)+len(s2_only)} items',
            delta_color='normal' if balanced else 'inverse',
        )

    # Totals row
    t1, t2, t3, t4 = st.columns(4)
    t1.metric('Recons Balanced', f'{sum(1 for r in results if abs(float(r[1].iloc[6]["Amount (₹)"] or 0)) < AMOUNT_TOLERANCE)} / 6')
    t2.metric('Total S1-Only Items', sum(len(r[2]) for r in results))
    t3.metric('Total S2-Only Items', sum(len(r[3]) for r in results))
    t4.metric('Total Discrepancies', sum(int(r[4]["_discrepancy"].sum()) if len(r[4]) > 0 else 0 for r in results))

    st.divider()

    # ── Filters ─────────────────────────────────
    st.subheader('🔍 Filters')
    fc1, fc2 = st.columns([2, 3])
    with fc1:
        inv_search = st.text_input('Search Invoice Number', placeholder='Type invoice no. to search…')
    with fc2:
        # Build entity list from Sales buyer_name column
        s_cm = get_cm(sales_df)
        buyer_col = s_cm.get('buyer_name')
        if buyer_col and buyer_col in sales_df.columns:
            entities = sorted(sales_df[buyer_col].dropna().astype(str).unique().tolist())
            entity_filter = st.multiselect('Filter by Buyer Entity', entities, placeholder='All entities')
        else:
            entity_filter = []

    st.divider()

    # ── Recon Tabs ──────────────────────────────
    tab_labels = [r[0] for r in results]
    tabs = st.tabs(tab_labels)

    for tab, (name, brs, s1_only, s2_only, matched, cm1, cm2) in zip(tabs, results):
        s1n, s2n = name.split(' vs ')
        with tab:

            # BRS Table
            st.markdown(f'#### Reconciliation Statement — {name}')
            diff_val = float(brs.iloc[6]['Amount (₹)'] or 0)
            if abs(diff_val) < AMOUNT_TOLERANCE:
                st.success('✅ BRS is balanced — Sources reconcile.')
            else:
                st.error(f'⚠️ BRS has unreconciled difference of ₹{diff_val:,.2f} — investigate.')
            st.dataframe(style_brs(brs), use_container_width=True, hide_index=True)

            st.markdown('---')

            # Apply filters to s1_only and s2_only
            def apply_filters(df, cm):
                out = df.copy()
                if inv_search:
                    inv_col = cm.get('inv')
                    if inv_col and inv_col in out.columns:
                        out = out[out[inv_col].astype(str).str.upper().str.contains(inv_search.upper(), na=False)]
                if entity_filter:
                    buyer_col = cm.get('buyer_name')
                    if buyer_col and buyer_col in out.columns:
                        out = out[out[buyer_col].astype(str).isin(entity_filter)]
                return out

            s1f = apply_filters(s1_only, cm1)
            s2f = apply_filters(s2_only, cm2)
            disc_count = int(matched['_discrepancy'].sum()) if len(matched) > 0 else 0

            # Sub-tabs
            sub_labels = [
                f'In {s1n} Only ({len(s1f):,})',
                f'In {s2n} Only ({len(s2f):,})',
                f'Discrepancies ({disc_count:,})',
            ]
            sub1, sub2, sub3 = st.tabs(sub_labels)

            with sub1:
                if len(s1f) == 0:
                    st.success('Nothing here — all items reconciled.' if not inv_search and not entity_filter else 'No results match the filter.')
                else:
                    # Category filter
                    cats = ['All'] + sorted(s1f['_category'].dropna().unique().tolist())
                    sel = st.selectbox('Filter by Category', cats, key=f's1cat_{name}')
                    df_show = s1f if sel == 'All' else s1f[s1f['_category'] == sel]

                    # Category breakdown
                    cat_counts = s1f['_category'].value_counts()
                    cc = st.columns(min(len(cat_counts), 4))
                    for idx, (cat, cnt) in enumerate(cat_counts.items()):
                        cc[idx % 4].metric(cat[:40], cnt)
                    st.markdown('')

                    tot = to_num(df_show[cm1['total']]).sum() if cm1.get('total') and cm1['total'] in df_show.columns else 0
                    st.caption(f'Showing {len(df_show):,} invoices | Total: ₹{tot:,.2f}')
                    render_table(df_show, cm1, height=min(400, max(150, len(df_show)*35 + 50)), key=f's1_{name}')

            with sub2:
                if len(s2f) == 0:
                    st.success('Nothing here — all items reconciled.' if not inv_search and not entity_filter else 'No results match the filter.')
                else:
                    cats = ['All'] + sorted(s2f['_category'].dropna().unique().tolist())
                    sel = st.selectbox('Filter by Category', cats, key=f's2cat_{name}')
                    df_show = s2f if sel == 'All' else s2f[s2f['_category'] == sel]

                    cat_counts = s2f['_category'].value_counts()
                    cc = st.columns(min(len(cat_counts), 4))
                    for idx, (cat, cnt) in enumerate(cat_counts.items()):
                        cc[idx % 4].metric(cat[:40], cnt)
                    st.markdown('')

                    tot = to_num(df_show[cm2['total']]).sum() if cm2.get('total') and cm2['total'] in df_show.columns else 0
                    st.caption(f'Showing {len(df_show):,} invoices | Total: ₹{tot:,.2f}')
                    render_table(df_show, cm2, height=min(400, max(150, len(df_show)*35 + 50)), key=f's2_{name}')

            with sub3:
                if disc_count == 0:
                    st.success('No discrepancies — all matched invoices have consistent values.')
                else:
                    disc = matched[matched['_discrepancy'] | matched['_gstin_diff']].copy()
                    if inv_search:
                        disc = disc[disc['_inv'].str.upper().str.contains(inv_search.upper(), na=False)]
                    disc_cols = [c for c in ['_inv','_m_sg_s1','_m_sg_s2','_m_bg_s1','_m_bg_s2',
                                              '_m_total_s1','_m_total_s2','_amt_diff','_gstin_diff']
                                 if c in disc.columns]
                    rename = {
                        '_inv':'Invoice No.', '_m_sg_s1':f'Seller GSTIN ({s1n})',
                        '_m_sg_s2':f'Seller GSTIN ({s2n})', '_m_bg_s1':f'Buyer GSTIN ({s1n})',
                        '_m_bg_s2':f'Buyer GSTIN ({s2n})', '_m_total_s1':f'Total ({s1n}) ₹',
                        '_m_total_s2':f'Total ({s2n}) ₹', '_amt_diff':'Diff (₹)',
                        '_gstin_diff':'GSTIN Mismatch',
                    }
                    st.caption(f'{len(disc):,} discrepancies')
                    st.dataframe(
                        disc[disc_cols].rename(columns=rename),
                        use_container_width=True, height=400, hide_index=True
                    )

    # ── Download ────────────────────────────────
    st.divider()
    st.subheader('📥 Download Excel Output')
    st.caption('Full recon with BRS, categorised exceptions, and discrepancies — formatted for sharing.')
    with st.spinner('Building Excel…'):
        excel_buf = build_excel(results)
    st.download_button(
        label='Download Recon Output (.xlsx)',
        data=excel_buf,
        file_name=f'Ecosystem_Recon_{datetime.now().strftime("%b_%Y")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        use_container_width=True,
    )
    st.caption('Swiggy Instamart Finance | Ecosystem Purchase Recon Tool')

if __name__ == '__main__':
    main()
