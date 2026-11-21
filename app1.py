import io
import os
import calendar
from collections import defaultdict
from datetime import datetime, date
import requests
import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import portrait, A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet

# ==========================================
# 1. ページ基本設定 ＆ デザインCSS
# ==========================================
st.set_page_config(
    page_title="請求書発行システム（株式会社小林断截）",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    .stApp { background-color: #F8FAFC; }
    h1, h2, h3 { color: #0F172A !important; font-family: 'Helvetica Neue', Arial, sans-serif; }
    div.stButton > button:first-child {
        background-color: #1E3A8A; color: white; border-radius: 6px; border: none; font-weight: bold; padding: 0.5rem 1rem;
    }
    div.stButton > button:first-child:hover { background-color: #1E40AF; color: white; }
    div.stDownloadButton > button:first-child {
        background-color: #0F172A; color: white; border-radius: 6px; border: none; font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================
# 2. パスワード認証（Streamlit Secrets）
# ==========================================
APP_PASSWORD = ""
if "auth" in st.secrets:
    APP_PASSWORD = str(st.secrets["auth"].get("password", ""))
elif "APP_PASSWORD" in st.secrets:
    APP_PASSWORD = str(st.secrets["APP_PASSWORD"])

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔐 ログイン - 請求書発行システム")
    if not APP_PASSWORD:
        st.error("⚠️ `.streamlit/secrets.toml` に `[auth] password` が設定されていません。")
        st.stop()

    with st.form("login_form"):
        pw_input = st.text_input("パスワードを入力してください", type="password")
        submitted = st.form_submit_button("ログイン")
        if submitted:
            if pw_input == APP_PASSWORD:
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("❌ パスワードが正しくありません。")
    st.stop()

st.title("🖨️ 請求書発行システム（株式会社小林断截）")

# ==========================================
# 3. 日本語フォント登録処理（NotoSansJP-Regular.ttf）
# ==========================================
def setup_japanese_font():
    font_candidates = [
        "NotoSansJP-Regular.ttf",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "NotoSansJP-Regular.ttf"),
    ]
    for font_path in font_candidates:
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont("JPFont", font_path))
                return "JPFont"
            except Exception as e:
                st.warning(f"⚠️ NotoSansJP-Regular.ttf の登録に失敗しました: {e}")

    st.warning("⚠️ `NotoSansJP-Regular.ttf` が見つかりません。同じフォルダに配置してください。")
    return "Helvetica"

FONT_NAME = setup_japanese_font()

# ==========================================
# 4. Secrets設定 ＆ 定数
# ==========================================
if "pocket" in st.secrets:
    POCKET_API_KEY = str(st.secrets["pocket"].get("api_key", "")).strip()
    POCKET_BASE_URL = str(st.secrets["pocket"].get("base_url", "https://app060.at-pocket.com/seihon03_bb")).strip().rstrip("/")
else:
    POCKET_API_KEY = ""
    POCKET_BASE_URL = "https://app060.at-pocket.com/seihon03_bb"

APP_ID = 18
TAX_RATE = 0.10
DETAIL_ROW_HEIGHT = 1.0 * cm   # 明細1行の高さ: 1cm
REMARK_BOX_HEIGHT = 3.0 * cm   # 備考欄の高さ: 3cm

COL_WIDTHS = [
    1.9 * cm,  # 売上日（4文字幅相当）
    5.6 * cm,  # 名目（10文字幅相当）
    2.0 * cm,  # 数量（4文字幅相当）
    2.2 * cm,  # 単価（4文字幅相当）
    3.4 * cm,  # 金額（6文字幅相当）
    2.9 * cm,  # 受注番号（5文字幅相当）
]
TOTAL_DETAIL_WIDTH = sum(COL_WIDTHS)  # 18.0cm

ISSUER_INFO = (
    "<b>株式会社小林断截</b><br/>"
    "〒130-0021 墨田区緑2-11-18<br/>"
    "info@k-dan.co.jp www.k-dan.co.jp<br/>"
    "TEL:03-3634-8351 FAX:03-3633-0695"
)

# ==========================================
# 5. 便利関数
# ==========================================
def clean_val(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    if isinstance(val, list):
        return "、".join([clean_val(v) for v in val if clean_val(v)])
    s = str(val).strip()
    if s.lower() in ["none", "nan", "nat", "null"]:
        return ""
    return s

def parse_date_safe(val):
    s = clean_val(val)
    if not s:
        return None
    try:
        if isinstance(val, (date, datetime)):
            return val.date() if isinstance(val, datetime) else val
        return datetime.strptime(s[:10].replace("/", "-"), "%Y-%m-%d").date()
    except ValueError:
        return None

def format_mmdd(d_val):
    parsed = d_val if isinstance(d_val, date) else parse_date_safe(d_val)
    if not parsed:
        return ""
    return parsed.strftime("%m/%d")

def parse_number_safe(val):
    s = clean_val(val).replace(",", "").replace("￥", "").replace("¥", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0

def format_num_display(num_val, is_currency=False):
    if num_val is None:
        return ""
    if abs(num_val) < 1e-9:
        return ""
    if abs(num_val - round(num_val)) < 1e-9:
        formatted = f"{int(round(num_val)):,}"
    else:
        formatted = f"{num_val:,.2f}"
    return f"￥{formatted}" if is_currency else formatted

def is_keijo_checked(val):
    if isinstance(val, list):
        return any(clean_val(v) == "計上" for v in val)
    s = clean_val(val)
    if not s:
        return False
    parts = [p.strip() for p in s.replace("、", ",").split(",")]
    return "計上" in parts or s == "計上"

def extract_subform_rows(record):
    sub_rows = []
    for k, v in record.items():
        if isinstance(v, list) and len(v) > 0 and isinstance(v[0], dict):
            first_rec = v[0].get("record", v[0]) if isinstance(v[0], dict) else {}
            if isinstance(first_rec, dict) and any(str(fk).startswith("field-140") for fk in first_rec.keys()):
                for item in v:
                    rec_sub = item.get("record", item) if isinstance(item, dict) else {}
                    if isinstance(rec_sub, dict):
                        sub_rows.append(rec_sub)
    return sub_rows

# ==========================================
# 6. @pocket API データ取得関数
# ==========================================
def fetch_pocket_records(start_date_str, end_date_str):
    if not POCKET_API_KEY:
        st.error("❌ `.streamlit/secrets.toml` に `[pocket] api_key` が設定されていません。")
        return []

    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}

    query_str = f'field-91 in ("計上") and field-143 >= "{start_date_str}" and field-143 <= "{end_date_str}"'
    params = {"query": query_str}

    try:
        st.info(f"📡 @pocket（APP {APP_ID}）からデータを取得中（売上日: {start_date_str} ～ {end_date_str}）...")
        response = requests.get(url, headers=headers, params=params, timeout=30)

        if response.status_code != 200:
            response = requests.get(url, headers=headers, timeout=30)

        if response.status_code == 200:
            result = response.json()
            raw_items = result.get("records", result.get("Data", result.get("Records", [])))
            parsed_records = []
            for item in raw_items:
                if not isinstance(item, dict):
                    continue
                rec_body = item.get("record", {})
                if not isinstance(rec_body, dict):
                    rec_body = {}
                parsed_records.append({**item, **rec_body})
            st.success(f"✨ 取得成功: {len(parsed_records)} 件のレコードを受信しました。")
            return parsed_records
        else:
            st.error(f"❌ APIエラー ({response.status_code}): {response.text}")
    except Exception as e:
        st.error(f"🚨 通信例外エラー: {e}")
    return []

# ==========================================
# 7. PDF生成関数（高さ1cm×12行組、超過時は2ページ目以降へ、備考高さ3cm）
# ==========================================
def generate_customer_invoice_pdf(customer_name, main_count, detail_rows, bank_transfer_info, remark_text):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.2 * cm,
        bottomMargin=1.2 * cm,
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "InvTitle", parent=styles["Normal"], fontName=FONT_NAME,
        fontSize=18, leading=22, alignment=1, textColor=colors.HexColor("#0F172A")
    )
    meta_style = ParagraphStyle(
        "MetaStyle", parent=styles["Normal"], fontName=FONT_NAME,
        fontSize=9, leading=13, textColor=colors.HexColor("#1E293B")
    )
    meta_right = ParagraphStyle(
        "Meta