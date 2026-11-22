import calendar
from collections import defaultdict
from datetime import date, datetime
import io
import os
import zipfile

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, portrait
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
import requests
import streamlit as st

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
      os.path.join(
          os.path.dirname(os.path.abspath(__file__)), "NotoSansJP-Regular.ttf"
      ),
  ]
  for font_path in font_candidates:
    if os.path.exists(font_path):
      try:
        pdfmetrics.registerFont(TTFont("JPFont", font_path))
        return "JPFont"
      except Exception as e:
        st.warning(f"⚠️ NotoSansJP-Regular.ttf の登録に失敗しました: {e}")

  st.warning(
      "⚠️ `NotoSansJP-Regular.ttf` が見つかりません。同じフォルダに配置してください。"
  )
  return "Helvetica"


FONT_NAME = setup_japanese_font()

# ==========================================
# 4. Secrets設定 ＆ 定数
# ==========================================
if "pocket" in st.secrets:
  POCKET_API_KEY = str(st.secrets["pocket"].get("api_key", "")).strip()
  POCKET_BASE_URL = (
      str(
          st.secrets["pocket"].get(
              "base_url", "https://app060.at-pocket.com/seihon03_bb"
          )
      )
      .strip()
      .rstrip("/")
  )
else:
  POCKET_API_KEY = ""
  POCKET_BASE_URL = "https://app060.at-pocket.com/seihon03_bb"

APP_ID = 18
TAX_RATE = 0.10
ROWS_PER_PAGE = 12
TITLE_SPACER_HEIGHT = 3.0 * cm  # タイトル下の段落余白: 3cm
BANK_BOX_HEIGHT = 3.0 * cm  # 振込先の高さ: 3cm
DETAIL_ROW_HEIGHT = 1.0 * cm  # 明細1行の高さ: 1cm
REMARK_BOX_HEIGHT = 3.0 * cm  # 備考欄の高さ: 3cm

TOTAL_WEIGHT = 3 + 16 + 4 + 3 + 6 + 4
TOTAL_DETAIL_WIDTH = 18.0 * cm
COL_WIDTHS = [
    TOTAL_DETAIL_WIDTH * (3 / TOTAL_WEIGHT),  # 売上日（3文字幅）
    TOTAL_DETAIL_WIDTH * (16 / TOTAL_WEIGHT),  # 名目（16文字幅）
    TOTAL_DETAIL_WIDTH * (4 / TOTAL_WEIGHT),  # 数量（4文字幅）
    TOTAL_DETAIL_WIDTH * (3 / TOTAL_WEIGHT),  # 単価（3文字幅）
    TOTAL_DETAIL_WIDTH * (6 / TOTAL_WEIGHT),  # 金額（6文字幅）
    TOTAL_DETAIL_WIDTH * (4 / TOTAL_WEIGHT),  # 受注番号（4文字幅）
]

ISSUER_COMPANY_NAME = "<b>株式会社小林断截</b>"
ISSUER_REG_NUMBER = "T9-0106-0101-0677"
DEFAULT_BANK_TRANSFER = "りそな銀行 本所支店（当座）0502479"


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


def sanitize_filename_part(name_str):
  cleaned = clean_val(name_str)
  for ch in ["/", "\\", ":", "*", "?", '"', "<", ">", "|"]:
    cleaned = cleaned.replace(ch, "_")
  return cleaned or "得意先未設定"


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


def has_valid_amount(val):
  s = clean_val(val).replace(",", "").replace("￥", "").replace("¥", "")
  if not s:
    return False
  try:
    float(s)
    return True
  except ValueError:
    return False


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
      first_rec = (
          v[0].get("record", v[0]) if isinstance(v[0], dict) else {}
      )
      if isinstance(first_rec, dict) and any(
          str(fk).startswith("field-140") for fk in first_rec.keys()
      ):
        for item in v:
          rec_sub = (
              item.get("record", item) if isinstance(item, dict) else {}
          )
          if isinstance(rec_sub, dict):
            sub_rows.append(rec_sub)
  return sub_rows


# ==========================================
# 6. @pocket API データ取得関数
# ==========================================
def fetch_pocket_records(start_date_str, end_date_str):
  if not POCKET_API_KEY:
    st.error(
        "❌ `.streamlit/secrets.toml` に `[pocket] api_key` が設定されていません。"
    )
    return []

  url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
  headers = {
      "X-At-Pocket-API-Key": POCKET_API_KEY,
      "Accept": "application/json",
  }

  query_str = f'field-91 in ("計上") and field-143 >= "{start_date_str}" and field-143 <= "{end_date_str}"'
  params = {"query": query_str}

  try:
    st.info(
        f"📡 @pocket（APP {APP_ID}）からデータを取得中（売上日: {start_date_str} ～"
        f" {end_date_str}）..."
    )
    response = requests.get(url, headers=headers, params=params, timeout=30)

    if response.status_code != 200:
      response = requests.get(url, headers=headers, timeout=30)

    if response.status_code == 200:
      result = response.json()
      raw_items = result.get(
          "records", result.get("Data", result.get("Records", []))
      )
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
# 7. Excel生成関数（顧客別売上一覧表＋全明細一覧）
# ==========================================
def generate_sales_summary_excel(
    grouped_data, remarks_dict, start_date_obj, end_date_obj
):
  wb = Workbook()

  font_title = Font(name="Meiryo UI", size=14, bold=True, color="0F172A")
  font_sub = Font(name="Meiryo UI", size=9.5, color="334155")
  font_header = Font(name="Meiryo UI", size=10, bold=True, color="FFFFFF")
  font_body = Font(name="Meiryo UI", size=10, color="0F172A")
  font_bold = Font(name="Meiryo UI", size=10, bold=True, color="0F172A")

  fill_header = PatternFill(
      start_color="1E3A8A", end_color="1E3A8A", fill_type="solid"
  )
  fill_zebra = PatternFill(
      start_color="F8FAFC", end_color="F8FAFC", fill_type="solid"
  )
  fill_total = PatternFill(
      start_color="DBEAFE", end_color="DBEAFE", fill_type="solid"
  )

  thin_side = Side(style="thin", color="CBD5E1")
  med_navy = Side(style="medium", color="1E3A8A")
  double_bottom = Side(style="double", color="1E3A8A")

  border_cell = Border(
      left=thin_side, right=thin_side, top=thin_side, bottom=thin_side
  )
  border_header = Border(
      left=thin_side, right=thin_side, top=med_navy, bottom=med_navy
  )
  border_total = Border(
      left=thin_side, right=thin_side, top=med_navy, bottom=double_bottom
  )

  period_label = (
      f"対象売上日：{start_date_obj.strftime('%Y年%m月%d日')} ～"
      f" {end_date_obj.strftime('%Y年%m月%d日')} ／"
      f" 出力日：{date.today().strftime('%Y年%m月%d日')}"
  )

  # ------------------------------------------
  # シート1：顧客別売上一覧
  # ------------------------------------------
  ws1 = wb.active
  ws1.title = "顧客別売上一覧"
  ws1.views.sheetView[0].showGridLines = True

  ws1["A1"] = "顧客別 売上一覧表（株式会社小林断截）"
  ws1["A1"].font = font_title
  ws1.row_dimensions[1].height = 26

  ws1["A2"] = period_label
  ws1["A2"].font = font_sub
  ws1.row_dimensions[2].height = 18

  headers_s1 = [
      "No.",
      "得意先名",
      "メイン件数",
      "明細行数",
      "税抜金額",
      "消費税（10%）",
      "合計金額（税込）",
      "備考",
  ]
  header_row_s1 = 4
  ws1.row_dimensions[header_row_s1].height = 24

  for col_idx, h_text in enumerate(headers_s1, start=1):
    cell = ws1.cell(row=header_row_s1, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.border = border_header
    cell.alignment = Alignment(horizontal="center", vertical="center")

  sorted_customers = sorted(grouped_data.keys())
  curr_row = header_row_s1 + 1

  for idx, cust_name in enumerate(sorted_customers, start=1):
    payload = grouped_data[cust_name]
    main_cnt = payload["main_count"]
    rows = payload["rows"]
    subtotal = int(round(sum(r["amount"] for r in rows)))
    tax = int(round(subtotal * TAX_RATE))
    total_inc = subtotal + tax
    remark_str = clean_val(remarks_dict.get(cust_name, ""))

    row_vals = [
        idx,
        cust_name,
        main_cnt,
        len(rows),
        subtotal,
        tax,
        total_inc,
        remark_str,
    ]
    ws1.row_dimensions[curr_row].height = 22
    is_even = idx % 2 == 0

    for col_idx, val in enumerate(row_vals, start=1):
      c = ws1.cell(row=curr_row, column=col_idx, value=val)
      c.font = font_body
      c.border = border_cell
      if is_even:
        c.fill = fill_zebra

      if col_idx == 1:
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.number_format = "#,##0"
      elif col_idx in (2, 8):
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
      elif col_idx in (3, 4):
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = "#,##0"
      else:
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.number_format = "¥#,##0"

    curr_row += 1

  first_data_row = header_row_s1 + 1
  last_data_row = curr_row - 1
  ws1.row_dimensions[curr_row].height = 25

  total_cells = {
      1: "",
      2: "合 計",
      3: (
          f"=SUM(C{first_data_row}:C{last_data_row})"
          if last_data_row >= first_data_row
          else 0
      ),
      4: (
          f"=SUM(D{first_data_row}:D{last_data_row})"
          if last_data_row >= first_data_row
          else 0
      ),
      5: (
          f"=SUM(E{first_data_row}:E{last_data_row})"
          if last_data_row >= first_data_row
          else 0
      ),
      6: (
          f"=SUM(F{first_data_row}:F{last_data_row})"
          if last_data_row >= first_data_row
          else 0
      ),
      7: (
          f"=SUM(G{first_data_row}:G{last_data_row})"
          if last_data_row >= first_data_row
          else 0
      ),
      8: "",
  }
  for col_idx, val in total_cells.items():
    c = ws1.cell(row=curr_row, column=col_idx, value=val)
    c.font = font_bold
    c.fill = fill_total
    c.border = border_total
    if col_idx == 2:
      c.alignment = Alignment(horizontal="center", vertical="center")
    elif col_idx in (3, 4):
      c.alignment = Alignment(horizontal="right", vertical="center")
      c.number_format = "#,##0"
    elif col_idx in (5, 6, 7):
      c.alignment = Alignment(horizontal="right", vertical="center")
      c.number_format = "¥#,##0"
    else:
      c.alignment = Alignment(horizontal="left", vertical="center")

  ws1.freeze_panes = f"A{header_row_s1 + 1}"
  ws1.auto_filter.ref = f"A{header_row_s1}:H{max(header_row_s1, last_data_row)}"

  col_widths_s1 = [7, 32, 13, 13, 16, 16, 18, 36]
  for col_idx, w in enumerate(col_widths_s1, start=1):
    ws1.column_dimensions[get_column_letter(col_idx)].width = w

  # ------------------------------------------
  # シート2：顧客別明細一覧
  # ------------------------------------------
  ws2 = wb.create_sheet(title="顧客別明細一覧")
  ws2.views.sheetView[0].showGridLines = True

  ws2["A1"] = "顧客別 売上明細一覧表（株式会社小林断截）"
  ws2["A1"].font = font_title
  ws2.row_dimensions[1].height = 26

  ws2["A2"] = period_label
  ws2["A2"].font = font_sub
  ws2.row_dimensions[2].height = 18

  headers_s2 = [
      "得意先名",
      "売上日",
      "名目",
      "数量",
      "単価",
      "金額",
      "受注番号",
      "区分",
  ]
  header_row_s2 = 4
  ws2.row_dimensions[header_row_s2].height = 24

  for col_idx, h_text in enumerate(headers_s2, start=1):
    cell = ws2.cell(row=header_row_s2, column=col_idx, value=h_text)
    cell.font = font_header
    cell.fill = fill_header
    cell.border = border_header
    cell.alignment = Alignment(horizontal="center", vertical="center")

  curr_row_2 = header_row_s2 + 1
  row_counter_2 = 0
  for cust_name in sorted_customers:
    for r in grouped_data[cust_name]["rows"]:
      row_counter_2 += 1
      ws2.row_dimensions[curr_row_2].height = 21
      is_even = row_counter_2 % 2 == 0
      vals_2 = [
          cust_name,
          r["sale_date_str"],
          r["item_name"],
          r["qty"],
          r["unit_price"],
          r["amount"],
          r["order_no"],
          r["source"],
      ]
      for col_idx, val in enumerate(vals_2, start=1):
        c = ws2.cell(row=curr_row_2, column=col_idx, value=val)
        c.font = font_body
        c.border = border_cell
        if is_even:
          c.fill = fill_zebra

        if col_idx in (1, 3):
          c.alignment = Alignment(horizontal="left", vertical="center")
        elif col_idx in (2, 7, 8):
          c.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx == 4:
          c.alignment = Alignment(horizontal="right", vertical="center")
          c.number_format = "#,##0"
        else:
          c.alignment = Alignment(horizontal="right", vertical="center")
          c.number_format = "¥#,##0"

      curr_row_2 += 1

  first_detail_row = header_row_s2 + 1
  last_detail_row = curr_row_2 - 1
  ws2.row_dimensions[curr_row_2].height = 25
  total_cells_2 = {
      1: "合 計",
      2: "",
      3: "",
      4: (
          f"=SUM(D{first_detail_row}:D{last_detail_row})"
          if last_detail_row >= first_detail_row
          else 0
      ),
      5: "",
      6: (
          f"=SUM(F{first_detail_row}:F{last_detail_row})"
          if last_detail_row >= first_detail_row
          else 0
      ),
      7: "",
      8: "",
  }
  for col_idx, val in total_cells_2.items():
    c = ws2.cell(row=curr_row_2, column=col_idx, value=val)
    c.font = font_bold
    c.fill = fill_total
    c.border = border_total
    if col_idx == 1:
      c.alignment = Alignment(horizontal="center", vertical="center")
    elif col_idx == 4:
      c.alignment = Alignment(horizontal="right", vertical="center")
      c.number_format = "#,##0"
    elif col_idx == 6:
      c.alignment = Alignment(horizontal="right", vertical="center")
      c.number_format = "¥#,##0"

  ws2.freeze_panes = f"A{header_row_s2 + 1}"
  ws2.auto_filter.ref = f"A{header_row_s2}:H{max(header_row_s2, last_detail_row)}"

  col_widths_s2 = [30, 11, 36, 12, 14, 16, 14, 10]
  for col_idx, w in enumerate(col_widths_s2, start=1):
    ws2.column_dimensions[get_column_letter(col_idx)].width = w

  out_buf = io.BytesIO()
  wb.save(out_buf)
  out_buf.seek(0)
  return out_buf


# ==========================================
# 8. PDF生成関数（洗練されたモダンヘッダー・インボイス登録番号対応）
# ==========================================
def generate_customer_invoice_pdf(
    customer_name, main_count, detail_rows, bank_transfer_info, remark_text
):
  buffer = io.BytesIO()
  doc = SimpleDocTemplate(
      buffer,
      pagesize=portrait(A4),
      leftMargin=1.5 * cm,
      rightMargin=1.5 * cm,
      topMargin=0.5 * cm,
      bottomMargin=0.5 * cm,
  )
  styles = getSampleStyleSheet()

  ink_dark = colors.HexColor("#0F172A")
  ink_muted = colors.HexColor("#475569")
  navy_primary = colors.HexColor("#1E3A8A")
  navy_accent = colors.HexColor("#2563EB")
  bg_header = colors.HexColor("#F1F5F9")
  bg_soft_blue = colors.HexColor("#EFF6FF")
  border_gray = colors.HexColor("#CBD5E1")

  title_main_style = ParagraphStyle(
      "InvTitleMain",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=19,
      leading=23,
      alignment=1,
      textColor=navy_primary,
  )
  title_sub_style = ParagraphStyle(
      "InvTitleSub",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=7.5,
      leading=9,
      alignment=1,
      textColor=ink_muted,
  )
  issuer_date_style = ParagraphStyle(
      "IssuerDateStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8.5,
      leading=11,
      alignment=2,
      textColor=ink_muted,
  )
  issuer_company_style = ParagraphStyle(
      "IssuerCompanyStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=14.5,
      leading=18,
      alignment=2,
      textColor=ink_dark,
  )
  issuer_addr_style = ParagraphStyle(
      "IssuerAddrStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=7.6,
      leading=10.2,
      alignment=2,
      textColor=ink_muted,
  )
  issuer_greeting_style = ParagraphStyle(
      "IssuerGreetingStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=7.6,
      leading=10.2,
      alignment=2,
      textColor=ink_dark,
  )
  cust_style = ParagraphStyle(
      "CustStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=13,
      leading=17,
      textColor=ink_dark,
  )
  meta_badge_style = ParagraphStyle(
      "MetaBadgeStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8.5,
      leading=11,
      textColor=navy_primary,
  )
  sum_label_style = ParagraphStyle(
      "SumLabel",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=7.8,
      leading=10,
      alignment=1,
      textColor=ink_muted,
  )
  sum_label_main_style = ParagraphStyle(
      "SumLabelMain",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8,
      leading=10,
      alignment=1,
      textColor=colors.white,
  )
  sum_val_main_style = ParagraphStyle(
      "SumValMain",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=12,
      leading=14,
      alignment=2,
      textColor=navy_primary,
  )
  sum_val_sub_style = ParagraphStyle(
      "SumValSub",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=9.5,
      leading=12,
      alignment=2,
      textColor=ink_dark,
  )
  th_style = ParagraphStyle(
      "THStyle",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8,
      leading=10,
      alignment=1,
      textColor=colors.white,
  )
  td_left = ParagraphStyle(
      "TDLeft",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8,
      leading=10,
      alignment=0,
      textColor=ink_dark,
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
      fontSize=9,
      leading=11,
      alignment=2,
      textColor=navy_primary,
  )
  box_body_style = ParagraphStyle(
      "BoxBody",
      parent=styles["Normal"],
      fontName=FONT_NAME,
      fontSize=8.2,
      leading=11.5,
      textColor=ink_dark,
  )

  subtotal_amount = sum(row["amount"] for row in detail_rows)
  tax_amount = int(round(subtotal_amount * TAX_RATE))
  grand_total_inc_tax = int(round(subtotal_amount + tax_amount))
  subtotal_ex_tax = int(round(subtotal_amount))

  chunks = [
      detail_rows[i : i + ROWS_PER_PAGE]
      for i in range(0, len(detail_rows), ROWS_PER_PAGE)
  ]
  if not chunks:
    chunks = [[]]

  elements = []
  issue_date_str = date.today().strftime("%Y年%m月%d日")
  transfer_clean = clean_val(bank_transfer_info).replace("\n", "<br/>")
  remark_clean = clean_val(remark_text).replace("\n", "<br/>")

  issuer_addr_html = (
      f"登録番号：{ISSUER_REG_NUMBER}<br/>"
      "〒130-0021 東京都墨田区緑2-11-18<br/>"
      "TEL：03-3634-8351 ／ FAX：03-3633-0695<br/>"
      "E-mail：info@k-dan.co.jp ／ URL：www.k-dan.co.jp"
  )
  issuer_greeting_html = (
      "下記の通りご請求申し上げます。<br/>"
      "期日までに下記口座へお振込みお願い申し上げます。"
  )

  for page_idx, chunk in enumerate(chunks):
    if page_idx > 0:
      elements.append(PageBreak())

    page_suffix = (
        f"（{page_idx + 1} / {len(chunks)}）" if len(chunks) > 1 else ""
    )

    title_tbl = Table(
        [
            [
                Paragraph(
                    f"<b>ご 請 求 書</b> <font size='10'>{page_suffix}</font>",
                    title_main_style,
                )
            ],
            [Paragraph("I N V O I C E", title_sub_style)],
        ],
        colWidths=[TOTAL_DETAIL_WIDTH],
    )
    title_tbl.setStyle(
        TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LINEBELOW", (0, 1), (0, 1), 1.5, navy_primary),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (0, 0), 1),
            ("BOTTOMPADDING", (0, 1), (0, 1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ])
    )
    elements.append(title_tbl)
    elements.append(Spacer(1, TITLE_SPACER_HEIGHT))

    left_cust_block = Table(
        [
            [
                Paragraph(
                    f"<b>{customer_name}</b> <font size='10'"
                    " color='#475569'>御中</font>",
                    cust_style,
                )
            ],
            [Paragraph(f"<b>対象件数：{main_count} 件</b>", meta_badge_style)],
        ],
        colWidths=[8.8 * cm],
    )
    left_cust_block.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LINEBELOW", (0, 0), (0, 0), 1.5, navy_primary),
            ("BACKGROUND", (0, 1), (0, 1), bg_soft_blue),
            ("LINEBEFORE", (0, 1), (0, 1), 3.0, navy_accent),
            ("TOPPADDING", (0, 0), (0, 0), 0),
            ("BOTTOMPADDING", (0, 0), (0, 0), 4),
            ("TOPPADDING", (0, 1), (0, 1), 4),
            ("BOTTOMPADDING", (0, 1), (0, 1), 4),
            ("LEFTPADDING", (0, 0), (0, 0), 0),
            ("LEFTPADDING", (0, 1), (0, 1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ])
    )

    right_issuer_block = Table(
        [
            [Paragraph(f"請求日：{issue_date_str}", issuer_date_style)],
            [Paragraph(ISSUER_COMPANY_NAME, issuer_company_style)],
            [Paragraph(issuer_addr_html, issuer_addr_style)],
            [Paragraph(issuer_greeting_html, issuer_greeting_style)],
        ],
        colWidths=[TOTAL_DETAIL_WIDTH - 9.2 * cm],
    )
    right_issuer_block.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (0, 0), 2),
            ("BOTTOMPADDING", (0, 1), (0, 1), 2),
            ("BOTTOMPADDING", (0, 2), (0, 2), 2.5),
            ("BOTTOMPADDING", (0, 3), (0, 3), 0),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ])
    )

    top_header_tbl = Table(
        [[left_cust_block, "", right_issuer_block]],
        colWidths=[8.8 * cm, 0.4 * cm, TOTAL_DETAIL_WIDTH - 9.2 * cm],
    )
    top_header_tbl.setStyle(
        TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])
    )
    elements.append(top_header_tbl)
    elements.append(Spacer(1, 0.15 * cm))

    summary_box = Table(
        [
            [
                Paragraph("<b>ご請求金額（税込）</b>", sum_label_main_style),
                Paragraph("<b>税抜金額</b>", sum_label_style),
                Paragraph("<b>消費税（10％）</b>", sum_label_style),
            ],
            [
                Paragraph(f"<b>￥{grand_total_inc_tax:,}</b>", sum_val_main_style),
                Paragraph(f"<b>￥{subtotal_ex_tax:,}</b>", sum_val_sub_style),
                Paragraph(f"<b>￥{tax_amount:,}</b>", sum_val_sub_style),
            ],
        ],
        colWidths=[6.6 * cm, 5.7 * cm, 5.7 * cm],
        rowHeights=[0.45 * cm, 0.65 * cm],
    )
    summary_box.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (0, 0), navy_primary),
            ("BACKGROUND", (1, 0), (2, 0), bg_header),
            ("BACKGROUND", (0, 1), (0, 1), bg_soft_blue),
            ("BACKGROUND", (1, 1), (2, 1), colors.white),
            ("BOX", (0, 0), (-1, -1), 1.2, navy_primary),
            ("LINEAFTER", (0, 0), (0, -1), 1.2, navy_primary),
            ("LINEAFTER", (1, 0), (1, -1), 0.6, border_gray),
            ("LINEBELOW", (1, 0), (2, 0), 0.6, border_gray),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ])
    )
    elements.append(summary_box)
    elements.append(Spacer(1, 0.12 * cm))

    transfer_box = Table(
        [[
            Paragraph(
                f"<font color='#1E3A8A'><b>■ お振込先口座</b></font><br/>{transfer_clean}",
                box_body_style,
            )
        ]],
        colWidths=[TOTAL_DETAIL_WIDTH],
        rowHeights=[BANK_BOX_HEIGHT],
    )
    transfer_box.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ("BOX", (0, 0), (-1, -1), 0.8, border_gray),
            ("LINEBEFORE", (0, 0), (0, -1), 3.0, navy_primary),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ])
    )
    elements.append(transfer_box)
    elements.append(Spacer(1, 0.12 * cm))

    detail_table_data = [[
        Paragraph("<b>売上日</b>", th_style),
        Paragraph("<b>名目</b>", th_style),
        Paragraph("<b>数量</b>", th_style),
        Paragraph("<b>単価</b>", th_style),
        Paragraph("<b>金額</b>", th_style),
        Paragraph("<b>受注番号</b>", th_style),
    ]]
    row_heights = [0.5 * cm]

    page_subtotal = 0.0
    for idx in range(ROWS_PER_PAGE):
      if idx < len(chunk):
        r = chunk[idx]
        page_subtotal += r["amount"]
        detail_table_data.append([
            Paragraph(r["sale_date_str"], td_center),
            Paragraph(r["item_name"], td_left),
            Paragraph(
                format_num_display(r["qty"], is_currency=False), td_right
            ),
            Paragraph(
                format_num_display(r["unit_price"], is_currency=True), td_right
            ),
            Paragraph(
                format_num_display(r["amount"], is_currency=True), td_right
            ),
            Paragraph(r["order_no"], td_center),
        ])
      else:
        detail_table_data.append(["", "", "", "", "", ""])
      row_heights.append(DETAIL_ROW_HEIGHT)

    detail_table_data.append([
        Paragraph("<b>小計（ページ計）</b>", td_center),
        Paragraph("", td_left),
        Paragraph("", td_right),
        Paragraph("", td_right),
        Paragraph(f"<b>￥{int(round(page_subtotal)):,}</b>", total_row_right),
        Paragraph("", td_center),
    ])
    row_heights.append(0.6 * cm)

    detail_tbl = Table(
        detail_table_data, colWidths=COL_WIDTHS, rowHeights=row_heights
    )
    last_row_idx = len(detail_table_data) - 1
    detail_tbl.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), navy_primary),
            ("BACKGROUND", (0, last_row_idx), (-1, last_row_idx), bg_soft_blue),
            ("GRID", (0, 1), (-1, last_row_idx - 1), 0.5, border_gray),
            ("LINEBEFORE", (0, 0), (-1, 0), 0.5, colors.HexColor("#3B82F6")),
            ("BOX", (0, 0), (-1, -1), 1.2, navy_primary),
            ("LINEABOVE", (0, last_row_idx), (-1, last_row_idx), 1.2, navy_primary),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 0),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
        ])
    )
    elements.append(detail_tbl)
    elements.append(Spacer(1, 0.12 * cm))

    remark_tbl = Table(
        [[
            Paragraph(
                f"<font color='#1E3A8A'><b>■ 備考</b></font><br/>{remark_clean}",
                box_body_style,
            )
        ]],
        colWidths=[TOTAL_DETAIL_WIDTH],
        rowHeights=[REMARK_BOX_HEIGHT],
    )
    remark_tbl.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("BOX", (0, 0), (-1, -1), 0.8, border_gray),
            ("LINEBEFORE", (0, 0), (0, -1), 3.0, border_gray),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ])
    )
    elements.append(remark_tbl)

  doc.build(elements)
  buffer.seek(0)
  return buffer


# ==========================================
# 8-2. 全得意先分 請求書PDF 一括ZIP生成関数
# ==========================================
def generate_all_invoices_zip(grouped_data, remarks_dict, bank_transfer_info):
  zip_buffer = io.BytesIO()
  issue_date_str = date.today().strftime("%Y%m%d")
  used_filenames = set()

  with zipfile.ZipFile(
      zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED
  ) as zf:
    for cust_name in sorted(grouped_data.keys()):
      payload = grouped_data[cust_name]
      main_cnt = payload["main_count"]
      rows = payload["rows"]
      remark_str = remarks_dict.get(cust_name, "")

      pdf_buf = generate_customer_invoice_pdf(
          customer_name=cust_name,
          main_count=main_cnt,
          detail_rows=rows,
          bank_transfer_info=bank_transfer_info,
          remark_text=remark_str,
      )

      safe_name = sanitize_filename_part(cust_name)
      base_filename = f"ご請求書_{safe_name}_{issue_date_str}.pdf"
      final_filename = base_filename
      dup_idx = 2
      while final_filename in used_filenames:
        final_filename = f"ご請求書_{safe_name}_{issue_date_str}_{dup_idx}.pdf"
        dup_idx += 1
      used_filenames.add(final_filename)

      zf.writestr(final_filename, pdf_buf.getvalue())

  zip_buffer.seek(0)
  return zip_buffer


# ==========================================
# 9. メイン画面
# ==========================================
today_date = date.today()
auto_start = date(today_date.year, today_date.month, 1)
auto_last_day = calendar.monthrange(today_date.year, today_date.month)[1]
auto_end = date(today_date.year, today_date.month, auto_last_day)

st.markdown("### 🔎 検索条件（売上計上＝計上・売上日期間）")
c1, c2, c3 = st.columns([2, 2, 2])
with c1:
  start_date_input = st.date_input(
      "売上日（開始・field-143）", value=auto_start, key="search_start_date"
  )
with c2:
  end_date_input = st.date_input(
      "売上日（終了・field-143）", value=auto_end, key="search_end_date"
  )
with c3:
  st.markdown("<br/>", unsafe_allow_html=True)
  fetch_clicked = st.button(
      "📥 @pocketからデータ取得・集計", type="primary", key="btn_fetch_pocket"
  )

if "input_bank_transfer" not in st.session_state:
  st.session_state["input_bank_transfer"] = DEFAULT_BANK_TRANSFER

st.markdown("### 🏦 振込先（UI入力・PDF枠高さ3cm）")
bank_transfer_input = st.text_area(
    "振込先口座情報を入力してください",
    height=75,
    key="input_bank_transfer",
)

if fetch_clicked:
  raw_records = fetch_pocket_records(
      start_date_input.strftime("%Y-%m-%d"), end_date_input.strftime("%Y-%m-%d")
  )

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
    main_amt_raw = rec.get("field-59")
    if main_name and has_valid_amount(main_amt_raw):
      main_amt = parse_number_safe(main_amt_raw)
      grouped_data[customer]["rows"].append({
          "sale_date_obj": main_date,
          "sale_date_str": format_mmdd(main_date),
          "item_name": main_name,
          "qty": main_qty,
          "unit_price": main_price,
          "amount": main_amt,
          "order_no": order_no,
          "source": "メイン",
      })

    sub_items = extract_subform_rows(rec)
    for sub in sub_items:
      part_9 = clean_val(sub.get("field-140_9"))
      part_3 = clean_val(sub.get("field-140_3"))
      merged_name = "，".join([p for p in [part_9, part_3] if p])

      sub_amt_raw = sub.get("field-140_7")
      if not merged_name or not has_valid_amount(sub_amt_raw):
        continue

      sub_date = parse_date_safe(sub.get("field-140_1")) or main_date
      sub_qty = parse_number_safe(sub.get("field-140_4"))
      sub_price = parse_number_safe(sub.get("field-140_5"))
      sub_amt = parse_number_safe(sub_amt_raw)

      grouped_data[customer]["rows"].append({
          "sale_date_obj": sub_date,
          "sale_date_str": format_mmdd(sub_date),
          "item_name": merged_name,
          "qty": sub_qty,
          "unit_price": sub_price,
          "amount": sub_amt,
          "order_no": order_no,
          "source": "サブ",
      })

  # 得意先ごとに明細を売上日順（同日は取得順）に並び替え
  for cust_name, payload in grouped_data.items():
    payload["rows"] = [
        r
        for _, r in sorted(
            enumerate(payload["rows"]),
            key=lambda pair: (
                pair[1].get("sale_date_obj") or date.max,
                pair[0],
            ),
        )
    ]

  st.session_state["grouped_invoice_data"] = dict(grouped_data)

if "grouped_invoice_data" in st.session_state:
  grouped = st.session_state["grouped_invoice_data"]
  if not grouped:
    st.warning(
        "⚠️ 条件（売上計上＝計上・指定期間内）に一致するレコードが見つかりませんでした。"
    )
  else:
    if "customer_remarks_18" not in st.session_state:
      st.session_state["customer_remarks_18"] = {}

    for cust_name in grouped.keys():
      if cust_name not in st.session_state["customer_remarks_18"]:
        st.session_state["customer_remarks_18"][cust_name] = ""

    current_period_key = f"{start_date_input}_{end_date_input}_{len(grouped)}"
    if st.session_state.get("zip_period_key") != current_period_key:
      st.session_state["zip_period_key"] = current_period_key
      st.session_state.pop("all_invoices_zip_bytes", None)
      st.session_state.pop("all_invoices_zip_filename", None)

    # ------------------------------------------
    # 一括出力セクション（売上一覧Excel ＆ 全請求書PDF ZIP）
    # ------------------------------------------
    st.markdown("### 📊 一括出力（売上一覧Excel ／ 全請求書PDF ZIP）")
    col_ex, col_zip_prep, col_zip_dl = st.columns([2.2, 1.8, 2.0])

    with col_ex:
      excel_bytes = generate_sales_summary_excel(
          grouped_data=grouped,
          remarks_dict=st.session_state["customer_remarks_18"],
          start_date_obj=start_date_input,
          end_date_obj=end_date_input,
      )
      st.download_button(
          label="📥 顧客別売上一覧表（.xlsx）",
          data=excel_bytes,
          file_name=(
              f"顧客別売上一覧表_{start_date_input.strftime('%Y%m%d')}-{end_date_input.strftime('%Y%m%d')}.xlsx"
          ),
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          key="dl_sales_summary_excel",
      )

    with col_zip_prep:
      if st.button(
          f"📦 一括ZIPを作成（{len(grouped)}社分）", key="btn_build_all_zip"
      ):
        with st.spinner("全得意先分の請求書PDFを生成中..."):
          zip_buf = generate_all_invoices_zip(
              grouped_data=grouped,
              remarks_dict=st.session_state["customer_remarks_18"],
              bank_transfer_info=bank_transfer_input,
          )
          st.session_state["all_invoices_zip_bytes"] = zip_buf.getvalue()
          st.session_state["all_invoices_zip_filename"] = (
              f"ご請求書一括_{start_date_input.strftime('%Y%m%d')}-{end_date_input.strftime('%Y%m%d')}.zip"
          )
        st.success("✅ 一括ZIPの準備が完了しました。")

    with col_zip_dl:
      if st.session_state.get("all_invoices_zip_bytes"):
        st.download_button(
            label="⬇️ 一括ZIPをダウンロード",
            data=st.session_state["all_invoices_zip_bytes"],
            file_name=st.session_state.get(
                "all_invoices_zip_filename", "ご請求書一括.zip"
            ),
            mime="application/zip",
            key="dl_all_invoices_zip",
        )
      else:
        st.caption("※「一括ZIPを作成」を押すとダウンロードボタンが表示されます")

    st.markdown("---")

    # ------------------------------------------
    # 得意先別 個別請求書PDF出力セクション
    # ------------------------------------------
    st.markdown("### 📄 得意先別 集計・備考入力・個別PDF出力")
    for cust_name, payload in grouped.items():
      main_cnt = payload["main_count"]
      rows = payload["rows"]
      subtotal = sum(r["amount"] for r in rows)
      tax = int(round(subtotal * TAX_RATE))
      total_inc = int(round(subtotal + tax))
      page_cnt = max(1, (len(rows) + ROWS_PER_PAGE - 1) // ROWS_PER_PAGE)

      with st.container():
        col_info, col_rem, col_dl = st.columns([3, 4, 2])
        with col_info:
          st.markdown(f"**🏢 {cust_name}**")
          st.write(
              f"メイン件数: **{main_cnt} 件** / 明細行数: **{len(rows)} 行**（{page_cnt}"
              " ページ）"
          )
          st.write(
              f"税抜: ￥{int(round(subtotal)):,} / 税: ￥{tax:,} / **税込:"
              f" ￥{total_inc:,}**"
          )
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
          safe_fname = sanitize_filename_part(cust_name)
          st.download_button(
              label="📥 個別PDF出力",
              data=pdf_bytes,
              file_name=f"ご請求書_{safe_fname}_{date.today().strftime('%Y%m%d')}.pdf",
              mime="application/pdf",
              key=f"dl_pdf_{cust_name}",
          )

      with st.expander(f"明細プレビュー（{cust_name}：{len(rows)}行）"):
        preview_df = pd.DataFrame([
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
        ])
        st.dataframe(preview_df, use_container_width=True)

      st.markdown("---")