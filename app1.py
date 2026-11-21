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
    h1, h2, h3 { color: #000000 !important; font-family: 'Helvetica Neue', Arial, sans-serif; }
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
ROWS_PER_PAGE = 14
DETAIL_ROW_HEIGHT = 1.0 * cm   # 明細1行の高さ: 1cm
REMARK_BOX_HEIGHT = 3.0 * cm   # 備考欄の高さ: 3cm

# 文字幅指定（売上日3・名目16・数量4・単価3・金額6・受注番号4 = 合計36文字相当）に合わせた列幅（合計 18.0cm）
TOTAL_WEIGHT = 3 + 16 + 4 + 3 + 6 + 4
TOTAL_DETAIL_WIDTH = 18.0 * cm
COL_WIDTHS = [
    TOTAL_DETAIL_WIDTH * (3 / TOTAL_WEIGHT),   # 売上日（3文字幅）
    TOTAL_DETAIL_WIDTH * (16 / TOTAL_WEIGHT),  # 名目（16文字幅）
    TOTAL_DETAIL_WIDTH * (4 / TOTAL_WEIGHT),   # 数量（4文字幅）
    TOTAL_DETAIL_WIDTH * (3 / TOTAL_WEIGHT),   # 単価（3文字幅）
    TOTAL_DETAIL_WIDTH * (6 / TOTAL_WEIGHT),   # 金額（6文字幅）
    TOTAL_DETAIL_WIDTH * (4 / TOTAL_WEIGHT),   # 受注番号（4文字幅）
]

ISSUER_INFO = (
    "<b>株式会社小林断截</b> 〒130-0021墨田区緑2-11-18<br/>"
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
# 7. PDF生成関数（太線・黒文字・高さ1cm×14行組・備考高さ3cm）
# ==========================================
def generate_customer_invoice_pdf(customer_name, main_count, detail_rows, bank_transfer_info, remark_text):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=portrait(A4),
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=0.8 * cm,
        bottomMargin=0.8 * cm,
    )
    styles = getSampleStyleSheet()
    ink = colors.black

    title_style = ParagraphStyle(
        "InvTitle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=18,
        leading=22,
        alignment=1,
        textColor=ink,
    )
    meta_style = ParagraphStyle(
        "MetaStyle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.5,
        leading=12,
        textColor=ink,
    )
    meta_right = ParagraphStyle(
        "MetaRight",
        parent=meta_style,
        alignment=2,
    )
    cust_style = ParagraphStyle(
        "CustStyle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=13,
        leading=17,
        textColor=ink,
    )
    sum_label_style = ParagraphStyle(
        "SumLabel",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.5,
        leading=11,
        alignment=1,
        textColor=ink,
    )
    sum_val_right = ParagraphStyle(
        "SumValRight",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=10,
        leading=13,
        alignment=2,
        textColor=ink,
    )
    th_style = ParagraphStyle(
        "THStyle",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.5,
        leading=10,
        alignment=1,
        textColor=ink,
    )
    td_left = ParagraphStyle(
        "TDLeft",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8,
        leading=10,
        alignment=0,
        textColor=ink,
    )
    td_center = ParagraphStyle(
        "TDCenter",
        parent=td_left,
        alignment=1,
    )
    td_right = ParagraphStyle(
        "TDRight",
        parent=td_left,
        alignment=2,
    )
    total_row_right = ParagraphStyle(
        "TotalRowRight",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.5,
        leading=11,
        alignment=2,
        textColor=ink,
    )
    remark_body_style = ParagraphStyle(
        "RemarkBody",
        parent=styles["Normal"],
        fontName=FONT_NAME,
        fontSize=8.5,
        leading=12,
        textColor=ink,
    )

    # 金額計算（消費税10%）
    subtotal_amount = sum(row["amount"] for row in detail_rows)
    tax_amount = int(round(subtotal_amount * TAX_RATE))
    grand_total_inc_tax = int(round(subtotal_amount + tax_amount))
    subtotal_ex_tax = int(round(subtotal_amount))

    # 14行ずつ分割（14行超過時は2ページ目以降へ繰り越し）
    chunks = [detail_rows[i:i + ROWS_PER_PAGE] for i in range(0, len(detail_rows), ROWS_PER_PAGE)]
    if not chunks:
        chunks = [[]]

    elements = []
    issue_date_str = date.today().strftime("%Y年%m月%d日")
    transfer_clean = clean_val(bank_transfer_info).replace("\n", "<br/>")
    remark_clean = clean_val(remark_text).replace("\n", "<br/>")

    for page_idx, chunk in enumerate(chunks):
        if page_idx > 0:
            elements.append(PageBreak())

        page_suffix = f" （{page_idx + 1}/{len(chunks)}）" if len(chunks) > 1 else ""
        elements.append(Paragraph(f"<b>ご請求書{page_suffix}</b>", title_style))
        elements.append(Spacer(1, 0.15 * cm))

        # ヘッダー上部：得意先名（下にライン）・件数 vs 請求日・発行元
        top_header_data = [
            [
                Paragraph(f"<b>{customer_name} 御中</b>", cust_style),
                Paragraph(f"請求日：{issue_date_str}", meta_right),
            ],
            [
                Paragraph(f"<b>件数：{main_count} 件</b>", meta_style),
                Paragraph(ISSUER_INFO, meta_right),
            ],
        ]
        top_header_tbl = Table(top_header_data, colWidths=[9.5 * cm, TOTAL_DETAIL_WIDTH - 9.5 * cm])
        top_header_tbl.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LINEBELOW", (0, 0), (0, 0), 1.8, ink),
                    ("BOTTOMPADDING", (0, 0), (0, 0), 4),
                    ("TOPPADDING", (0, 1), (-1, 1), 3),
                ]
            )
        )
        elements.append(top_header_tbl)
        elements.append(Spacer(1, 0.18 * cm))

        # 集計ブロック（合計金額税込・税抜金額・消費税（10%）を右寄せ、太枠）
        summary_box = Table(
            [
                [
                    Paragraph("<b>合計金額税込</b>", sum_label_style),
                    Paragraph("<b>税抜金額</b>", sum_label_style),
                    Paragraph("<b>消費税（10％）</b>", sum_label_style),
                ],
                [
                    Paragraph(f"<b>￥{grand_total_inc_tax:,}</b>", sum_val_right),
                    Paragraph(f"<b>￥{subtotal_ex_tax:,}</b>", sum_val_right),
                    Paragraph(f"<b>￥{tax_amount:,}</b>", sum_val_right),
                ],
            ],
            colWidths=[6.0 * cm, 6.0 * cm, 6.0 * cm],
            rowHeights=[0.48 * cm, 0.65 * cm],
        )
        summary_box.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                    ("GRID", (0, 0), (-1, -1), 1.0, ink),
                    ("BOX", (0, 0), (-1, -1), 1.6, ink),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(summary_box)
        elements.append(Spacer(1, 0.15 * cm))

        # 振込先ブロック（UI入力）
        transfer_box = Table(
            [[Paragraph(f"<b>振込先：</b> {transfer_clean if transfer_clean else '（未入力）'}", meta_style)]],
            colWidths=[TOTAL_DETAIL_WIDTH],
        )
        transfer_box.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 1.4, ink),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(transfer_box)
        elements.append(Spacer(1, 0.2 * cm))

        # 明細テーブル（ヘッダー + 高さ1cm×14行組 + 金額縦合計行）
        detail_table_data = [
            [
                Paragraph("<b>売上日</b>", th_style),
                Paragraph("<b>名目</b>", th_style),
                Paragraph("<b>数量</b>", th_style),
                Paragraph("<b>単価</b>", th_style),
                Paragraph("<b>金額</b>", th_style),
                Paragraph("<b>受注番号</b>", th_style),
            ]
        ]
        row_heights = [0.55 * cm]

        page_subtotal = 0.0
        for idx in range(ROWS_PER_PAGE):
            if idx < len(chunk):
                r = chunk[idx]
                page_subtotal += r["amount"]
                detail_table_data.append(
                    [
                        Paragraph(r["sale_date_str"], td_center),
                        Paragraph(r["item_name"], td_left),
                        Paragraph(format_num_display(r["qty"], is_currency=False), td_right),
                        Paragraph(format_num_display(r["unit_price"], is_currency=True), td_right),
                        Paragraph(format_num_display(r["amount"], is_currency=True), td_right),
                        Paragraph(r["order_no"], td_center),
                    ]
                )
            else:
                detail_table_data.append(["", "", "", "", "", ""])
            row_heights.append(DETAIL_ROW_HEIGHT)

        # 下段：金額縦合計（自動計算）
        detail_table_data.append(
            [
                Paragraph("<b>合計</b>", td_center),
                Paragraph("", td_left),
                Paragraph("", td_right),
                Paragraph("", td_right),
                Paragraph(f"<b>￥{int(round(page_subtotal)):,}</b>", total_row_right),
                Paragraph("", td_center),
            ]
        )
        row_heights.append(0.65 * cm)

        detail_tbl = Table(detail_table_data, colWidths=COL_WIDTHS, rowHeights=row_heights)
        last_row_idx = len(detail_table_data) - 1
        detail_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E5E7EB")),
                    ("BACKGROUND", (0, last_row_idx), (-1, last_row_idx), colors.HexColor("#F3F4F6")),
                    ("GRID", (0, 0), (-1, -1), 1.0, ink),
                    ("BOX", (0, 0), (-1, -1), 1.8, ink),
                    ("LINEBELOW", (0, 0), (-1, 0), 1.6, ink),
                    ("LINEABOVE", (0, last_row_idx), (-1, last_row_idx), 1.6, ink),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 1),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                ]
            )
        )
        elements.append(detail_tbl)
        elements.append(Spacer(1, 0.2 * cm))

        # フッター：備考（UI入力・高さ3cm・幅は明細幅と同じ）
        remark_content = Paragraph(
            f"<b>備考</b><br/>{remark_clean}",
            remark_body_style,
        )
        remark_tbl = Table(
            [[remark_content]],
            colWidths=[TOTAL_DETAIL_WIDTH],
            rowHeights=[REMARK_BOX_HEIGHT],
        )
        remark_tbl.setStyle(
            TableStyle(
                [
                    ("BOX", (0, 0), (-1, -1), 1.6, ink),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        elements.append(remark_tbl)

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 8. メイン画面
# ==========================================
today_date = date.today()
auto_start = date(today_date.year, today_date.month, 1)
auto_last_day = calendar.monthrange(today_date.year, today_date.month)[1]
auto_end = date(today_date.year, today_date.month, auto_last_day)

st.markdown("### 🔎 検索条件（売上計上＝計上・売上日期間）")
c1, c2, c3 = st.columns([2, 2, 2])
with c1:
    start_date_input = st.date_input("売上日（開始・field-143）", value=auto_start, key="search_start_date")
with c2:
    end_date_input = st.date_input("売上日（終了・field-143）", value=auto_end, key="search_end_date")
with c3:
    st.markdown("<br/>", unsafe_allow_html=True)
    fetch_clicked = st.button("📥 @pocketからデータ取得・集計", type="primary", key="btn_fetch_pocket")

st.markdown("### 🏦 振込先（UI入力・全請求書共通）")
bank_transfer_input = st.text_area(
    "振込先口座情報を入力してください",
    value="〇〇銀行 〇〇支店（普通）1234567\n口座名義：カ）コバヤシダンサイ",
    height=75,
    key="input_bank_transfer",
)

if fetch_clicked:
    raw_records = fetch_pocket_records(start_date_input.strftime("%Y-%m-%d"), end_date_input.strftime("%Y-%m-%d"))

    grouped_data = defaultdict(lambda: {"main_count": 0, "rows": []})

    for rec in raw_records:
        if not is_keijo_checked(rec.get("field-91")):
            continue

        main_date = parse_date_safe(rec.get("field-143"))
        if not main_date or not (start_date_input <= main_date <= end_date_input):
            continue

        customer = clean_val(rec.get("field18-84")) or "（得意先名未設定）"
        order_no = clean_val(rec.get("js25"))

        grouped_data[customer]["main_count"] += 1

        main_name = clean_val(rec.get("field18-15"))
        main_qty = parse_number_safe(rec.get("field18-18"))
        main_price = parse_number_safe(rec.get("field-20"))
        main_amt_raw = clean_val(rec.get("field-59"))
        main_amt = parse_number_safe(main_amt_raw) if main_amt_raw else (main_qty * main_price)

        grouped_data[customer]["rows"].append(
            {
                "sale_date_str": format_mmdd(main_date),
                "item_name": main_name,
                "qty": main_qty,
                "unit_price": main_price,
                "amount": main_amt,
                "order_no": order_no,
                "source": "メイン",
            }
        )

        sub_items = extract_subform_rows(rec)
        for sub in sub_items:
            sub_date = parse_date_safe(sub.get("field-140_1")) or main_date
            part_9 = clean_val(sub.get("field-140_9"))
            part_3 = clean_val(sub.get("field-140_3"))
            merged_name = "，".join([p for p in [part_9, part_3] if p])

            sub_qty = parse_number_safe(sub.get("field-140_4"))
            sub_price = parse_number_safe(sub.get("field-140_5"))
            sub_amt_raw = clean_val(sub.get("field-140_7"))
            sub_amt = parse_number_safe(sub_amt_raw) if sub_amt_raw else (sub_qty * sub_price)

            grouped_data[customer]["rows"].append(
                {
                    "sale_date_str": format_mmdd(sub_date),
                    "item_name": merged_name,
                    "qty": sub_qty,
                    "unit_price": sub_price,
                    "amount": sub_amt,
                    "order_no": order_no,
                    "source": "サブ",
                }
            )

    st.session_state["grouped_invoice_data"] = dict(grouped_data)

if "grouped_invoice_data" in st.session_state:
    grouped = st.session_state["grouped_invoice_data"]
    if not grouped:
        st.warning("⚠️ 条件（売上計上＝計上・指定期間内）に一致するレコードが見つかりませんでした。")
    else:
        if "customer_remarks_18" not in st.session_state:
            st.session_state["customer_remarks_18"] = {}

        st.markdown("### 📄 得意先別 集計・備考入力・PDF出力")
        for cust_name, payload in grouped.items():
            main_cnt = payload["main_count"]
            rows = payload["rows"]
            subtotal = sum(r["amount"] for r in rows)
            tax = int(round(subtotal * TAX_RATE))
            total_inc = int(round(subtotal + tax))
            page_cnt = max(1, (len(rows) + ROWS_PER_PAGE - 1) // ROWS_PER_PAGE)

            if cust_name not in st.session_state["customer_remarks_18"]:
                st.session_state["customer_remarks_18"][cust_name] = ""

            with st.container():
                col_info, col_rem, col_dl = st.columns([3, 4, 2])
                with col_info:
                    st.markdown(f"**🏢 {cust_name}**")
                    st.write(f"メイン件数: **{main_cnt} 件** / 明細行数: **{len(rows)} 行**（{page_cnt} ページ）")
                    st.write(f"税抜: ￥{int(round(subtotal)):,} / 税: ￥{tax:,} / **税込: ￥{total_inc:,}**")
                with col_rem:
                    st.session_state["customer_remarks_18"][cust_name] = st.text_area(
                        f"備考（高さ3cm枠・{cust_name}）",
                        value=st.session_state["customer_remarks_18"][cust_name],
                        key=f"remark_area_{cust_name}",
                        height=75,
                    )
                with col_dl:
                    st.markdown("<br/>", unsafe_allow_html=True)
                    pdf_bytes = generate_customer_invoice_pdf(
                        customer_name=cust_name,
                        main_count=main_cnt,
                        detail_rows=rows,
                        bank_transfer_info=bank_transfer_input,
                        remark_text=st.session_state["customer_remarks_18"][cust_name],
                    )
                    safe_fname = cust_name.replace("/", "_").replace("\\", "_")
                    st.download_button(
                        label="📥 PDF出力",
                        data=pdf_bytes,
                        file_name=f"ご請求書_{safe_fname}_{date.today().strftime('%Y%m%d')}.pdf",
                        mime="application/pdf",
                        key=f"dl_pdf_{cust_name}",
                    )

            with st.expander(f"明細プレビュー（{cust_name}：{len(rows)}行）"):
                preview_df = pd.DataFrame(
                    [
                        {
                            "売上日": r["sale_date_str"],
                            "名目": r["item_name"],
                            "数量": r["qty"],
                            "単価": r["unit_price"],
                            "金額": r["amount"],
                            "受注番号": r["order_no"],
                            "区分": r["source"],
                        }
                        for r in rows
                    ]
                )
                st.dataframe(preview_df, use_container_width=True)

            st.markdown("---")