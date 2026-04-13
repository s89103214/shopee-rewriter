"""
蝦皮標題批次改寫 — 核心改寫邏輯
同步自 keyword-rewriter五家店.html（2026-04-13）
"""

import os
import re
import random
from openai import OpenAI
from config import (
    GIFT_MALE, GIFT_FEMALE, GIFT_NEUTRAL, CAT_GIFT,
    MALE_SIGNALS, FEMALE_SIGNALS, SHOPEE_RULE, MOMO_RULE,
    CAT_SIGNALS, CAT_KW_HINTS, ALL_GIFT_WORDS
)

# 取得 API Key：優先用 Streamlit Cloud secrets，其次用 .env
def _get_api_key():
    try:
        import streamlit as st
        return st.secrets.get('QWEN_API_KEY')
    except Exception:
        pass
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except Exception:
        pass
    return os.getenv('QWEN_API_KEY')

client = OpenAI(
    api_key=_get_api_key(),
    base_url='https://dashscope-intl.aliyuncs.com/compatible-mode/v1'
)


def detect_gender(text: str) -> str:
    """偵測商品性別屬性"""
    t = text.lower()
    m = sum(1 for w in MALE_SIGNALS if w in t)
    f = sum(1 for w in FEMALE_SIGNALS if w in t)
    if m > f:
        return 'male'
    if f > m:
        return 'female'
    return 'neutral'


def detect_category(text: str) -> str:
    """偵測商品品類"""
    t = text.lower()
    best, best_score = '通用', 0
    for cat, signals in CAT_SIGNALS.items():
        score = sum(1 for w in signals if w.lower() in t)
        if score > best_score:
            best_score = score
            best = cat
    return best


def get_gift_pool(text: str, platform: str = 'shopee') -> list:
    """根據品類和性別取得送禮詞池"""
    gender = detect_gender(text)
    cat = detect_category(text)

    if cat in CAT_GIFT:
        cat_pool = CAT_GIFT[cat]
        # 有些品類用 shopee/momo 分（如宗教開運）
        if isinstance(cat_pool, dict):
            if platform in cat_pool:
                plat_data = cat_pool[platform]
                if isinstance(plat_data, list):
                    return plat_data
                if isinstance(plat_data, dict):
                    return plat_data.get(gender, plat_data.get('neutral', []))
            # 直接用性別分
            if gender in cat_pool:
                return cat_pool[gender]
            if 'neutral' in cat_pool:
                return cat_pool['neutral']
        elif isinstance(cat_pool, list):
            return cat_pool

    # 預設：通用性別池
    if gender == 'male':
        return GIFT_MALE
    elif gender == 'female':
        return GIFT_FEMALE
    return GIFT_NEUTRAL


def pick_gifts(text: str, used: list = None, platform: str = 'shopee') -> list:
    """隨機挑選 1~2 個送禮關鍵字"""
    n = 1 if random.random() < 0.5 else 2
    pool = get_gift_pool(text, platform)

    if used:
        avail = [g for g in pool if g not in used]
        if len(avail) < n:
            avail = pool
    else:
        avail = pool

    return random.sample(avail, min(n, len(avail)))


def get_cat_hint(cat: str) -> str:
    """取得品類搜尋維度提示"""
    if cat in CAT_KW_HINTS:
        hints = CAT_KW_HINTS[cat]
        lines = [f"  {k}：{v}" for k, v in hints.items()]
        return '\n\n【品類搜尋維度參考（替換次要詞時從這些維度選）】\n' + '\n'.join(lines)
    return ''


def rev6(text: str) -> str:
    """前 6 個詞反轉順序，增加差異化"""
    tokens = text.strip().split()
    if len(tokens) <= 1:
        return text
    n = min(6, len(tokens))
    return ' '.join(tokens[:n][::-1] + tokens[n:])


def clean_ai_result(text: str, trailing_code: str = '') -> str:
    """清理 AI 回傳結果"""
    res = text.strip()
    res = re.sub(r'^[「『"\']+|[」』"\']+$', '', res)
    if trailing_code:
        res = re.sub(r'\b' + re.escape(trailing_code) + r'\b', '', res)
        res = re.sub(r'\s{2,}', ' ', res).strip()
        res = re.sub(r'\s+\d{2,5}\s*$', '', res).strip()
    return res


def trim_to_limit(text: str, limit: int) -> str:
    """強制裁剪到字數上限，保護送禮詞不被砍"""
    if len(text) <= limit:
        return text
    tokens = text.split()
    # 先從尾巴砍非送禮詞
    for i in range(len(tokens) - 1, 0, -1):
        if len(' '.join(tokens)) <= limit:
            break
        if tokens[i] not in ALL_GIFT_WORDS:
            tokens.pop(i)
    # 如果還超過，只好砍送禮詞
    while len(tokens) > 1 and len(' '.join(tokens)) > limit:
        tokens.pop()
    return ' '.join(tokens)


def ensure_gifts(text: str, gift: list) -> str:
    """確保結果剛好有 1~2 個送禮詞（不多不少）"""
    tokens = text.split()
    existing = [w for w in tokens if w in ALL_GIFT_WORDS]
    no_gift = [w for w in tokens if w not in ALL_GIFT_WORDS]
    # 合併 pickGift 選的 + AI 自己生的，取前 2 個不重複
    candidates = list(gift) + existing
    final = []
    seen = set()
    for g in candidates:
        if g not in seen and len(final) < 2:
            seen.add(g)
            final.append(g)
    return ' '.join(no_gift + final)


def split_trailing_code(title: str) -> tuple:
    """抽取標題尾巴的型號數字"""
    m = re.match(r'^(.*?)\s+(\d{2,5})\s*$', title)
    if m:
        return m.group(1).strip(), m.group(2)
    return title.strip(), ''


def build_rewrite_prompt(original_title: str, store_label: str,
                         history: list = None, platform: str = 'shopee',
                         trailing_code: str = '') -> dict:
    """
    建構改寫 prompt，回傳 {prompt, gift}
    """
    gift = pick_gifts(original_title, platform=platform)
    gender = detect_gender(original_title)

    if gender == 'male':
        gender_hint = '（男性商品：請用父親節禮物、送爸爸、送男友等，不要用母親節禮物）'
    elif gender == 'female':
        gender_hint = '（女性商品：請用母親節禮物、送媽媽、送女友等，不要用父親節禮物）'
    else:
        gender_hint = ''

    hist_text = ''
    if history:
        hist_lines = [f'版本{i+1}（{h["store"]}）：{h["text"]}' for i, h in enumerate(history)]
        hist_text = '\n\n【去重：次要關鍵字必須替換，不可與以下版本重複】\n' + '\n'.join(hist_lines)

    code_len = (len(trailing_code) + 1) if trailing_code else 0
    cat = detect_category(original_title)
    cat_hint = get_cat_hint(cat)
    cat_rule = ''
    if cat != '通用':
        cat_rule = f'\n6. 【品類鎖定：{cat}】所有關鍵字必須屬於「{cat}」品類，嚴禁出現其他品類的詞'

    code_note = ''
    if trailing_code:
        code_note = f'\n7. 【禁止】不要輸出尾巴的型號數字「{trailing_code}」，系統會自動補上'

    if platform == 'momo':
        plat_rule = MOMO_RULE
        plat_label = 'MOMO（模糊匹配）'
        char_limit = 50 - code_len
    else:
        plat_rule = SHOPEE_RULE
        plat_label = '蝦皮（精確匹配）'
        char_limit = 60 - code_len

    prompt = f"""你是電商店群關鍵字專家。

將「A店」關鍵字改寫成「{store_label}」的新版本。
A店平台：{plat_label}
{store_label}平台：{plat_label}

{plat_rule}

【多維度覆蓋策略】替換次要詞時，確保新詞覆蓋不同的搜尋維度（品項別名、材質、功能、對象、場合、風格），同一維度不重複。

改寫規則：
1. 【前6個詞不動】A 店的前 6 個關鍵字必須原封不動保留（系統會自動調換順序來製造差異），絕對不可替換、修改、合併或拆開這 6 個詞
2. 【結構不動】保持與 A 店相同的排列結構：主品項詞在前，屬性詞在中間，送禮詞分散在後半段
3. 只替換第 7 個詞以後的部分次要關鍵字，換成「不同維度」的詞（例：把場合詞換成風格詞，把功能詞換成對象詞）
4. 【字數要求】總字數（含空格）必須接近 {char_limit} 字（允許 ±2 字），不可太短也不可超過
5. 只輸出關鍵字，詞間空格，不加說明{cat_rule}{code_note}{cat_hint}{hist_text}

【A店（{len(original_title)}字）】
{original_title}

⚠️ 最後提醒：輸出裡【必須】包含以下送禮詞（原封不動放進去，不可省略、不可替換）：{'、'.join(gift)}{gender_hint}

請直接輸出{store_label}關鍵字（目標 {char_limit} 字）："""

    return {'prompt': prompt, 'gift': gift}


def rewrite_title(original_title: str, store_label: str, history: list = None,
                   platform: str = 'shopee') -> str:
    """呼叫 Qwen API 改寫一個標題"""
    keywords, code = split_trailing_code(original_title)

    result = build_rewrite_prompt(keywords, store_label, history,
                                   platform=platform, trailing_code=code)
    prompt = result['prompt']
    gift = result['gift']

    response = client.chat.completions.create(
        model='qwen-plus',
        max_tokens=500,
        messages=[{'role': 'user', 'content': prompt}]
    )

    raw = response.choices[0].message.content.strip()
    res = clean_ai_result(raw, code)

    # 確保剛好 1~2 個送禮詞
    res = ensure_gifts(res, gift)

    # 強制裁剪（送禮詞受保護）
    if platform == 'momo':
        kw_limit = 50
    else:
        kw_limit = 60 - (len(code) + 1 if code else 0)
    res = trim_to_limit(res, kw_limit)
    res = rev6(res)

    if code:
        res = res + ' ' + code
    return res
