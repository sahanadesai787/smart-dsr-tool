import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='Smart DSR Tool', page_icon='📊', layout='wide')
st.title('📊 Smart DSR Tool')
st.caption('Upload your daily sales order and master files → map automatically → download the DSR Excel.')

TARGET_COLS = ['Source','Ean_comb','TTL Vendor Cost','Channel','TTL QTY','TTL ESP','TTL MRP','Vendor Cost','TTL RGM','Cat 1','Cat 2','Cat 3','Cat 4','Style Code','Size','MRP','Myntra ID','Date','Time','Hour','Week','Location']

def read_upload(upload, preferred=None):
    if upload is None: return None
    name = upload.name.lower()
    try:
        if name.endswith('.csv'):
            return pd.read_csv(upload, dtype=object, low_memory=False)
        xls = pd.ExcelFile(upload)
        sheet = preferred if preferred in xls.sheet_names else xls.sheet_names[0]
        df = pd.read_excel(upload, sheet_name=sheet, dtype=object)
        # Find the real header if the first rows are title/blank rows.
        if all(str(c).startswith('Unnamed:') for c in df.columns):
            upload.seek(0)
            raw = pd.read_excel(upload, sheet_name=sheet, header=None, nrows=12, dtype=object)
            header_row = None
            for i, row in raw.iterrows():
                vals = [str(v).strip().lower() for v in row.tolist() if pd.notna(v)]
                if any(v in ['icode','item sku code','sale order item code','channel name'] for v in vals):
                    header_row = i; break
            if header_row is not None:
                upload.seek(0); df = pd.read_excel(upload, sheet_name=sheet, header=header_row, dtype=object)
        return df
    except Exception as e:
        st.error(f'Could not read {upload.name}: {e}')
        return None

def norm(s):
    return ''.join(ch.lower() for ch in str(s).strip() if ch.isalnum())

def find_col(df, *names):
    if df is None: return None
    lookup = {norm(c): c for c in df.columns}
    for name in names:
        if norm(name) in lookup: return lookup[norm(name)]
    # allow partial header matches
    for c in df.columns:
        nc = norm(c)
        for name in names:
            nn = norm(name)
            if nn and (nn in nc or nc in nn): return c
    return None

def series(df, col, default=''):
    if df is None or col is None: return pd.Series([default]*0)
    return df[col]

def clean_channel(value, channel_map):
    if pd.isna(value): return ''
    raw = str(value).strip()
    if channel_map is not None and len(channel_map):
        src = find_col(channel_map, 'Channel')
        dst = find_col(channel_map, 'Channel-Clean', 'Channel Clean')
        if src and dst:
            matches = channel_map[channel_map[src].astype(str).str.strip().str.lower() == raw.lower()]
            if not matches.empty and pd.notna(matches.iloc[0][dst]): return str(matches.iloc[0][dst]).strip()
    u = raw.upper()
    for key, label in [('MYNTRA','Myntra'),('AJIO','AJIO'),('FLIPKART','Flipkart'),('COCOBLU','COCOBLU'),('NYKAA','Nykaa'),('TATACLIQ','TataCliq'),('LIMEROAD','Limeroad'),('NOON','NOON'),('NAMSHI','NAMSHI'),('MAGENTO','Magento')]:
        if key in u: return label
    return raw

def make_output(sales, catalog, item_master, channels, vendor, weeks, myntra):
    if sales is None or sales.empty: raise ValueError('Sales Order Report is required and must contain rows.')
    sku_col = find_col(sales, 'Item SKU Code', 'Sale Order Item Code', 'SKU Code')
    itemcode_col = find_col(sales, 'Sale Order Item Code')
    channel_col = find_col(sales, 'Channel Name', 'Channel')
    price_col = find_col(sales, 'Selling Price', 'ESP', 'SellingPrice')
    mrp_col_sales = find_col(sales, 'MRP')
    date_col = find_col(sales, 'Order Date as dd/mm/yyyy hh:MM:ss', 'Order Date', 'Created Date')
    status_col = find_col(sales, 'Sale Order Item Status', 'Sale Order Status')
    qty_col = find_col(sales, 'Quantity', 'Qty', 'TTL QTY', 'Item Quantity')
    city_col = find_col(sales, 'Shipping Address City', 'Location', 'Facility')
    if sku_col is None: raise ValueError('Could not find Item SKU Code / Sale Order Item Code in sales report.')
    # Catalog lookup keys and relevant fields
    cat_sku = find_col(catalog, 'OLD SKU', 'Old SKU', 'SKU')
    cat_icode = find_col(catalog, 'ICODE', 'EAN', 'EAN_comb')
    cat_style = find_col(catalog, 'New Style', 'New Style Code', 'Style Code', 'New Style')
    cat_size = find_col(catalog, 'Size')
    cat_mrp = find_col(catalog, 'MRP')
    cat_cats = [find_col(catalog, 'C1', 'Cat 1'), find_col(catalog, 'C2', 'Cat 2'), find_col(catalog, 'C3', 'Cat 3'), find_col(catalog, 'C4', 'Cat 4')]
    cat_myntra = find_col(catalog, 'Style ID', 'Myntra ID', 'SKU ID')
    # item master mapping, used as fallback for item attributes/cost
    im_sku = find_col(item_master, 'OLD SKU', 'New SKU', 'SKU')
    im_icode = find_col(item_master, 'ICODE', 'EAN')
    im_style = find_col(item_master, 'New Style', 'New Style Code')
    im_cost = find_col(item_master, 'Vendor Cost', 'VC', 'FINAL VC')
    im_size = find_col(item_master, 'Size')
    im_cats = [find_col(item_master, 'C1','Cat 1'),find_col(item_master,'C2','Cat 2'),find_col(item_master,'C3','Cat 3'),find_col(item_master,'C4','Cat 4')]
    def lookup_map(df, keycol, valcol):
        if df is None or keycol is None or valcol is None: return {}
        temp = df[[keycol,valcol]].dropna(subset=[keycol]).copy()
        temp[keycol] = temp[keycol].astype(str).str.strip().str.lower()
        temp = temp.drop_duplicates(keycol, keep='last')
        return dict(zip(temp[keycol], temp[valcol]))
    cat_key_map = lookup_map(catalog, cat_sku, cat_icode)
    cat_style_map = lookup_map(catalog, cat_sku, cat_style)
    cat_size_map = lookup_map(catalog, cat_sku, cat_size)
    cat_mrp_map = lookup_map(catalog, cat_sku, cat_mrp)
    cat_cat_maps = [lookup_map(catalog, cat_sku, c) for c in cat_cats]
    cat_myntra_map = lookup_map(catalog, cat_style, cat_myntra)
    im_key_map = lookup_map(item_master, im_sku, im_icode)
    im_style_map = lookup_map(item_master, im_sku, im_style)
    im_cost_map = lookup_map(item_master, im_sku, im_cost)
    im_size_map = lookup_map(item_master, im_sku, im_size)
    im_cat_maps = [lookup_map(item_master, im_sku, c) for c in im_cats]
    # vendor cost can be keyed by new style/style/style new
    v_style = find_col(vendor, 'Style New', 'Style', 'New Style', 'Style Code')
    v_cost = find_col(vendor, 'FINAL VC', 'Vendor Cost', 'VC')
    vendor_map = lookup_map(vendor, v_style, v_cost)
    # dates/week map
    w_date = find_col(weeks, 'Date New', 'Date', 'Date.1', 'Date.2')
    w_week = find_col(weeks, 'G Week', 'week', 'week No', 'Week No')
    week_map = {}
    if weeks is not None and w_date and w_week:
        wd = pd.to_datetime(weeks[w_date], errors='coerce', dayfirst=True)
        for d, w in zip(wd, weeks[w_week]):
            if pd.notna(d): week_map[d.normalize()] = w
    # optional Myntra report mapping: infer style id and listing ID fields
    out = pd.DataFrame(index=sales.index)
    sku = sales[sku_col].astype(str).str.strip()
    sku_key = sku.str.lower()
    def mapped(primary, fallback=None, fallback_map=None):
        vals = sku_key.map(primary).where(lambda x: x.notna() & (x.astype(str)!='nan')) if primary else pd.Series(np.nan,index=sales.index)
        if fallback and fallback_map:
            vals = vals.fillna(sku_key.map(fallback_map))
        return vals
    ean = mapped(cat_key_map, im_key_map, im_key_map)
    style = mapped(cat_style_map, im_style_map, im_style_map)
    size = mapped(cat_size_map, im_size_map, im_size_map)
    mrp_catalog = mapped(cat_mrp_map)
    # fallback to sales MRP
    mrp = pd.to_numeric(mrp_catalog, errors='coerce').fillna(pd.to_numeric(sales[mrp_col_sales], errors='coerce') if mrp_col_sales else np.nan)
    vc_by_style = pd.Series(style, index=sales.index).astype(str).str.strip().str.lower().map(vendor_map)
    vc_by_sku = sku_key.map(im_cost_map)
    vc = pd.to_numeric(vc_by_style, errors='coerce').fillna(pd.to_numeric(vc_by_sku, errors='coerce'))
    channel_raw = sales[channel_col] if channel_col else pd.Series(['']*len(sales), index=sales.index)
    channel = channel_raw.map(lambda x: clean_channel(x, channels))
    qty = pd.to_numeric(sales[qty_col], errors='coerce').fillna(1) if qty_col else pd.Series(1,index=sales.index,dtype=float)
    qty = qty.fillna(1)
    price = pd.to_numeric(sales[price_col], errors='coerce').fillna(0) if price_col else pd.Series(0,index=sales.index,dtype=float)
    date_raw = pd.to_datetime(sales[date_col], errors='coerce', dayfirst=True) if date_col else pd.Series(pd.NaT,index=sales.index)
    out['Source'] = channel_raw.fillna('')
    out['Ean_comb'] = ean
    out['TTL Vendor Cost'] = vc * qty
    out['Channel'] = channel
    out['TTL QTY'] = qty
    # ESP per unit follows the provided rules: listed channels use selling price; others use selling price.
    # Total ESP is price x quantity, and formulas are embedded in the workbook after export.
    out['TTL ESP'] = price * qty
    out['TTL MRP'] = mrp * qty
    out['Vendor Cost'] = vc
    out['TTL RGM'] = out['TTL ESP'] - out['TTL Vendor Cost']
    for i, c in enumerate(['Cat 1','Cat 2','Cat 3','Cat 4']):
        out[c] = mapped(cat_cat_maps[i], im_cat_maps[i], im_cat_maps[i])
    out['Style Code'] = style
    out['Size'] = size
    out['MRP'] = mrp
    out['Myntra ID'] = pd.Series(style,index=sales.index).astype(str).str.strip().str.lower().map(cat_myntra_map)
    out['Date'] = date_raw.dt.strftime('%d/%m/%Y')
    out['Time'] = date_raw.dt.strftime('%H:%M:%S')
    out['Hour'] = date_raw.dt.hour
    out['Week'] = date_raw.dt.normalize().map(week_map)
    out['Location'] = sales[city_col] if city_col else ''
    # Preserve unmatched input values for troubleshooting
    out = out[TARGET_COLS]
    return out, sales, {'sku_col':sku_col,'status_col':status_col,'price_col':price_col,'date_col':date_col,'channel_col':channel_col,'qty_col':qty_col}

st.subheader('1. Upload files')
with st.expander('Required files', expanded=True):
    sales_file = st.file_uploader('Sales Order Report (daily)', type=['xlsx','xls','csv'], key='sales')
    c1,c2 = st.columns(2)
    catalog_file = c1.file_uploader('Globus Cataloging', type=['xlsx','xls','csv'], key='catalog')
    item_file = c2.file_uploader('Item Master (optional)', type=['xlsx','xls','csv'], key='item')
    c3,c4 = st.columns(2)
    channel_file = c3.file_uploader('Channel Master', type=['xlsx','xls','csv'], key='channel')
    vendor_file = c4.file_uploader('Vendor Cost / Season Change Bucket', type=['xlsx','xls','csv'], key='vendor')
    c5,c6 = st.columns(2)
    week_file = c5.file_uploader('Week Master', type=['xlsx','xls','csv'], key='week')
    myntra_file = c6.file_uploader('Myntra MTD Sales Report (optional)', type=['xlsx','xls','csv'], key='myntra')
    template_file = st.file_uploader('DSR Template (optional — keeps your format)', type=['xlsx','xls'], key='template')

st.info('Daily use: upload the newest sales order report and the latest master files, then click Generate DSR. Files are processed in this session only.')
if st.button('⚙️ Generate DSR Excel', type='primary', use_container_width=True):
    if not sales_file or not catalog_file:
        st.error('Please upload at least the Sales Order Report and Globus Cataloging file.')
    else:
        with st.spinner('Reading files and mapping records...'):
            sales = read_upload(sales_file)
            catalog = read_upload(catalog_file, 'Master')
            item = read_upload(item_file) if item_file else None
            channels = read_upload(channel_file) if channel_file else None
            vendor = read_upload(vendor_file) if vendor_file else None
            weeks = read_upload(week_file) if week_file else None
            myntra = read_upload(myntra_file) if myntra_file else None
            try:
                result, raw_sales, detected = make_output(sales,catalog,item,channels,vendor,weeks,myntra)
                unmatched = pd.DataFrame({
                    'Sales SKU': raw_sales[detected['sku_col']].astype(str),
                    'Ean_comb missing': result['Ean_comb'].isna() | result['Ean_comb'].astype(str).str.strip().isin(['','nan','None']),
                    'Style Code missing': result['Style Code'].isna() | result['Style Code'].astype(str).str.strip().isin(['','nan','None']),
                    'Vendor Cost missing': result['Vendor Cost'].isna(),
                    'Week missing': result['Week'].isna(),
                })
                issues = unmatched[unmatched.iloc[:,1:].any(axis=1)].copy()
                st.success(f'DSR generated: {len(result):,} rows.')
                a,b,c,d = st.columns(4)
                a.metric('Sales rows', f'{len(result):,}')
                b.metric('Missing EAN', f'{int(unmatched["Ean_comb missing"].sum()):,}')
                c.metric('Missing vendor cost', f'{int(unmatched["Vendor Cost missing"].sum()):,}')
                d.metric('Missing week', f'{int(unmatched["Week missing"].sum()):,}')
                st.subheader('DSR preview')
                st.dataframe(result.head(100), use_container_width=True, height=360)
                # Build workbook, optionally using the uploaded template's first sheet as the visible DSR layout.
                buf = BytesIO()
                with pd.ExcelWriter(buf, engine='openpyxl') as writer:
                    result.to_excel(writer, index=False, sheet_name='DSR')
                    if len(issues): issues.to_excel(writer, index=False, sheet_name='Mapping Issues')
                    pd.DataFrame([
                        ['TTL Vendor Cost','Vendor Cost × TTL QTY'],['TTL ESP','Selling Price × TTL QTY (source export values)'],['TTL MRP','MRP × TTL QTY'],['TTL RGM','TTL ESP − TTL Vendor Cost'],['Vendor Cost','Looked up by Style Code from Vendor Cost file; fallback to Item Master'],['Ean_comb','Looked up by Item SKU Code / OLD SKU in Cataloging, fallback to Item Master'],['Channel','Mapped using Channel Master'],['Cat 1–4','Looked up from Cataloging by SKU'],['Week','Looked up from Week Master using order date']
                    ],columns=['Field','Calculation / Mapping rule']).to_excel(writer,index=False,sheet_name='Formula Guide')
                buf.seek(0)
                wb = load_workbook(buf)
                ws = wb['DSR']
                # Style output and embed Excel formulas in calculation columns; calculated values remain available from app preview.
                headers = {cell.value: cell.column for cell in ws[1]}
                for cell in ws[1]:
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='244062')
                    cell.alignment = Alignment(wrap_text=True, vertical='center')
                ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
                for row in range(2, ws.max_row+1):
                    q = get_column_letter(headers['TTL QTY']); vc = get_column_letter(headers['Vendor Cost']); esp = get_column_letter(headers['TTL ESP']); mrpcol = get_column_letter(headers['MRP'])
                    ws.cell(row,headers['TTL Vendor Cost']).value = f'=IFERROR({vc}{row}*{q}{row},0)'
                    ws.cell(row,headers['TTL MRP']).value = f'=IFERROR({mrpcol}{row}*{q}{row},0)'
                    ws.cell(row,headers['TTL RGM']).value = f'=IFERROR({esp}{row}-{get_column_letter(headers["TTL Vendor Cost"])}{row},0)'
                for sheet in wb.worksheets:
                    sheet.freeze_panes = 'A2'
                    for col_cells in sheet.columns:
                        max_len = min(max((len(str(cell.value)) if cell.value is not None else 0 for cell in list(col_cells)[:250]), default=10)+2, 32)
                        sheet.column_dimensions[get_column_letter(col_cells[0].column)].width = max(12,max_len)
                final = BytesIO(); wb.save(final); final.seek(0)
                st.download_button('⬇️ Download DSR Excel', data=final.getvalue(), file_name=f'Smart_DSR_{datetime.now():%Y%m%d_%H%M}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
                if len(issues):
                    st.warning('Some mappings are missing. Open the Mapping Issues sheet and correct the master data before submitting the DSR.')
            except Exception as e:
                st.error(f'Could not generate DSR: {e}')

st.divider()
st.caption('Note: The tool does not assume order statuses should be excluded. Check your company DSR rules before submitting. Always review the Mapping Issues sheet and the ESP rule for each channel.')
