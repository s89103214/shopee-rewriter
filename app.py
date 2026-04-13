"""
蝦皮標題批次改寫 — Streamlit 介面
員工上傳蝦皮匯出的 Excel → 選擇要生成的店版本 → 一鍵改寫 → 下載結果
"""

import streamlit as st
import time
import zipfile
import xml.etree.ElementTree as ET
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


# ── Excel 寫入（保留原始格式，只改 C 欄內容） ──

def build_output_excel(original_bytes: bytes, products: list) -> bytes:
    """複製原始 Excel，在共用字串表新增新標題，只改 C 欄的索引值，格式完全保留"""
    bio_in = BytesIO(original_bytes)
    z_in = zipfile.ZipFile(bio_in, 'r')
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

    title_map = {p['row_num']: p['new_title'] for p in products if 'new_title' in p}

    # ── 1. 解析 sharedStrings.xml，把新標題加到尾巴 ──
    ss_tree = ET.parse(z_in.open('xl/sharedStrings.xml'))
    ss_root = ss_tree.getroot()

    # 計算目前有幾個共用字串
    existing_count = len(ss_root.findall(f'{{{ns}}}si'))

    # 建立 row_num → 新的共用字串索引 的對照表
    new_index_map = {}
    for row_num, new_title in title_map.items():
        idx = existing_count + len(new_index_map)
        new_index_map[row_num] = idx
        # 新增 <si><t>新標題</t></si>
        si_el = ET.SubElement(ss_root, f'{{{ns}}}si')
        t_el = ET.SubElement(si_el, f'{{{ns}}}t')
        t_el.text = new_title

    # 更新 count 和 uniqueCount 屬性
    total = existing_count + len(new_index_map)
    ss_root.set('count', str(total))
    ss_root.set('uniqueCount', str(total))

    # ── 2. 解析 sheet1.xml，只改 C 欄 cell 的 <v> 值 ──
    sheet_tree = ET.parse(z_in.open('xl/worksheets/sheet1.xml'))
    sheet_root = sheet_tree.getroot()

    for row in sheet_root.findall(f'.//{{{ns}}}row'):
        r_num = int(row.get('r'))
        if r_num not in new_index_map:
            continue

        for cell in row.findall(f'{{{ns}}}c'):
            ref = cell.get('r')
            if ref and ref.startswith('C'):
                # 保持 t="s"（共用字串類型），只改索引值
                v_el = cell.find(f'{{{ns}}}v')
                if v_el is not None:
                    v_el.text = str(new_index_map[r_num])
                break

    # ── 3. 重新打包 Excel（只替換 sheet1.xml 和 sharedStrings.xml） ──
    bio_out = BytesIO()
    with zipfile.ZipFile(bio_out, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.namelist():
            if item == 'xl/worksheets/sheet1.xml':
                z_out.writestr(item, ET.tostring(sheet_root, xml_declaration=True, encoding='UTF-8'))
            elif item == 'xl/sharedStrings.xml':
                z_out.writestr(item, ET.tostring(ss_root, xml_declaration=True, encoding='UTF-8'))
            else:
                z_out.writestr(item, z_in.read(item))

    z_in.close()
    return bio_out.getvalue()


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
    store_c = col2.checkbox('C店', value=False)
    store_d = col3.checkbox('D店', value=False)
    store_e = col4.checkbox('E店', value=False)

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
