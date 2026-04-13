"""
蝦皮標題批次改寫 — Streamlit 介面
員工上傳蝦皮匯出的 Excel → 選擇要生成的店版本 → 一鍵改寫 → 下載結果
"""

import streamlit as st
import time
import openpyxl
from copy import copy
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


# ── Excel 讀寫（用 openpyxl） ──

def read_shopee_excel(file_bytes: bytes) -> list:
    """讀取蝦皮匯出的 Excel，回傳商品列表"""
    bio = BytesIO(file_bytes)
    wb = openpyxl.load_workbook(bio)
    ws = wb.active

    products = []
    for row in ws.iter_rows(min_row=7, values_only=False):
        r_num = row[0].row
        product_id = str(row[0].value or '')
        sku = str(row[1].value or '')
        title = str(row[2].value or '')

        if title and title != 'None':
            products.append({
                'row_num': r_num,
                'product_id': product_id,
                'sku': sku,
                'title': title,
            })

    wb.close()
    return products


def build_output_excel(original_bytes: bytes, products: list) -> bytes:
    """複製原始 Excel，把 C 欄標題替換成改寫後的版本，回傳 bytes"""
    bio_in = BytesIO(original_bytes)
    wb = openpyxl.load_workbook(bio_in)
    ws = wb.active

    title_map = {p['row_num']: p['new_title'] for p in products if 'new_title' in p}

    for row_num, new_title in title_map.items():
        ws.cell(row=row_num, column=3, value=new_title)

    bio_out = BytesIO()
    wb.save(bio_out)
    wb.close()
    return bio_out.getvalue()


# ── 介面 ──

uploaded = st.file_uploader('上傳蝦皮匯出的 Excel 檔案', type=['xlsx'])

if uploaded:
    file_bytes = uploaded.read()

    # 初始化 session state
    if 'products' not in st.session_state or st.session_state.get('uploaded_name') != uploaded.name:
        st.session_state.products = read_shopee_excel(file_bytes)
        st.session_state.uploaded_name = uploaded.name
        st.session_state.file_bytes = file_bytes
        st.session_state.results = {}

    products = st.session_state.products
    st.success(f'找到 {len(products)} 個商品')

    # 顯示商品列表
    with st.expander('📋 商品列表（點擊展開）', expanded=False):
        for i, p in enumerate(products):
            st.text(f'{i+1}. [{p["sku"]}] {p["title"]}')

    st.divider()

    # 選擇要生成的店
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

    # 開始改寫
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

    # 顯示結果 & 下載
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

                # 下載按鈕
                output_bytes = build_output_excel(st.session_state.file_bytes, store_products)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f'rewritten_{store_label}_{timestamp}.xlsx'

                st.download_button(
                    label=f'⬇️ 下載 {store_label} Excel',
                    data=output_bytes,
                    file_name=filename,
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    key=f'dl_{store_label}'
                )
