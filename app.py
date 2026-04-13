"""
蝦皮標題批次改寫 — Streamlit 介面
員工上傳蝦皮匯出的 Excel → 選擇要生成的店版本 → 一鍵改寫 → 下載結果
"""

import streamlit as st
import time
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
from datetime import datetime
from io import BytesIO
from rewriter import rewrite_title

st.set_page_config(page_title='蝦皮關鍵字批次改寫', layout='wide')

# ── 密碼驗證 ──
if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title('🔒 請輸入密碼')
    pwd = st.text_input('密碼', type='password')
    if st.button('登入'):
        if pwd == st.secrets.get('APP_PASSWORD', 'mary8888'):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error('密碼錯誤')
    st.stop()

st.title('🔄 蝦皮關鍵字批次改寫工具')


# ── Excel 讀取（用 XML 直接解析，避免 openpyxl 相容性問題） ──

def read_shopee_excel(file_bytes: bytes) -> list:
    """讀取蝦皮匯出的 Excel"""
    bio = BytesIO(file_bytes)
    z = zipfile.ZipFile(bio)
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

    ss_tree = ET.parse(z.open('xl/sharedStrings.xml'))
    strings = []
    for si in ss_tree.findall(f'{ns}si'):
        text = ''.join(t.text or '' for t in si.iter(f'{ns}t'))
        strings.append(text)

    tree = ET.parse(z.open('xl/worksheets/sheet1.xml'))
    root = tree.getroot()

    products = []
    for row in root.findall(f'.//{ns}row'):
        r_num = int(row.get('r'))
        if r_num < 7:
            continue

        cells = {}
        for cell in row.findall(f'{ns}c'):
            ref = cell.get('r')
            col_letter = ''.join(c for c in ref if c.isalpha())
            t = cell.get('t')
            v_el = cell.find(f'{ns}v')
            v = v_el.text if v_el is not None else None
            if t == 's' and v is not None:
                val = strings[int(v)]
            else:
                val = v
            cells[col_letter] = val

        title = cells.get('C')
        if title:
            products.append({
                'row_num': r_num,
                'product_id': cells.get('A', ''),
                'sku': cells.get('B', ''),
                'title': title,
            })

    return products


# ── Excel 寫入（新建簡單 Excel，蝦皮批量匯入格式） ──

def build_output_excel(products: list) -> bytes:
    """建立新 Excel，包含蝦皮需要的欄位"""
    wb = openpyxl.Workbook()
    ws = wb.active

    # 蝦皮批量匯入的標頭
    ws.append(['商品ID', '主商品貨號', '商品名稱'])

    for p in products:
        ws.append([
            p.get('product_id', ''),
            p.get('sku', ''),
            p.get('new_title', p.get('title', ''))
        ])

    bio = BytesIO()
    wb.save(bio)
    wb.close()
    return bio.getvalue()


# ── 介面 ──

uploaded = st.file_uploader('上傳蝦皮匯出的 Excel 檔案', type=['xlsx'])

if uploaded:
    file_bytes = uploaded.read()

    if 'products' not in st.session_state or st.session_state.get('uploaded_name') != uploaded.name:
        st.session_state.products = read_shopee_excel(file_bytes)
        st.session_state.uploaded_name = uploaded.name
        st.session_state.file_bytes = file_bytes
        st.session_state.results = {}

    products = st.session_state.products
    st.success(f'找到 {len(products)} 個商品')

    with st.expander('📋 商品列表（點擊展開）', expanded=False):
        for i, p in enumerate(products):
            st.text(f'{i+1}. [{p["sku"]}] {p["title"]}')

    st.divider()

    st.subheader('選擇要生成的店版本')
    col1, col2, col3, col4 = st.columns(4)
    store_b = col1.checkbox('B店', value=True)
    store_c = col2.checkbox('C店', value=True)
    store_d = col3.checkbox('D店', value=True)
    store_e = col4.checkbox('E店', value=True)

    selected_stores = []
    if store_b: selected_stores.append('B')
    if store_c: selected_stores.append('C')
    if store_d: selected_stores.append('D')
    if store_e: selected_stores.append('E')

    platform = st.radio('平台', ['shopee', 'momo'], horizontal=True)

    st.divider()

    if st.button('🚀 開始批次改寫', type='primary', disabled=len(selected_stores) == 0):
        total = len(products) * len(selected_stores)
        progress = st.progress(0, text='準備中...')
        status = st.empty()
        done_count = 0

        results = {}

        for store in selected_stores:
            store_label = f'{store}店'
            status.info(f'正在改寫 {store_label}...')

            history = []
            store_products = []

            for idx, product in enumerate(products):
                title = product['title']
                progress.progress(
                    done_count / total,
                    text=f'{store_label} - {idx+1}/{len(products)}：{title[:30]}...'
                )

                try:
                    new_title = rewrite_title(title, store_label, history, platform=platform)
                    p_copy = dict(product)
                    p_copy['new_title'] = new_title
                    store_products.append(p_copy)
                    history.append({'store': store_label, 'text': new_title})
                except Exception as e:
                    p_copy = dict(product)
                    p_copy['new_title'] = title
                    p_copy['error'] = str(e)
                    store_products.append(p_copy)

                done_count += 1
                time.sleep(0.3)

            results[store_label] = store_products

        progress.progress(1.0, text='全部完成！')
        status.success('✅ 全部改寫完成！')
        st.session_state.results = results

    if st.session_state.results:
        st.divider()
        st.subheader('📥 改寫結果')

        for store_label, store_products in st.session_state.results.items():
            with st.expander(f'{store_label} 結果', expanded=True):
                for p in store_products:
                    err = p.get('error', '')
                    if err:
                        st.error(f'❌ [{p["sku"]}] 失敗：{err}')
                    else:
                        st.markdown(f'**原始：** {p["title"]}')
                        st.markdown(f'**{store_label}：** {p["new_title"]}')
                        st.text('')

                output_bytes = build_output_excel(store_products)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'rewritten_{store_label}_{timestamp}.xlsx'

                st.download_button(
                    label=f'⬇️ 下載 {store_label} Excel',
                    data=output_bytes,
                    file_name=filename,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=f'dl_{store_label}'
                )
