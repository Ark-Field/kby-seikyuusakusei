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
        if isinstance(val, (date,