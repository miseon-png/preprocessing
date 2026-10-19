import base64
import copy
import io
import os
import re
import sqlite3
import traceback
from datetime import datetime
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 및 원료코드 System", layout="wide")

DB_FILE = "stock_master.db"

# ---------------------------------------------------------
# DB 초기화 및 관리 함수 (전월 마감 롯트 데이터베이스)
# ---------------------------------------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS closed_lots (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item_name TEXT,
            yymmdd TEXT,
            avail_g REAL,
            vendor TEXT,
            closed_ym TEXT
        )
    """)
    conn.commit()
    conn.close()

def get_latest_closed_lots():
    """가장 최근 마감월의 잔여 롯트 목록을 가져옵니다."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT MAX(closed_ym) FROM closed_lots")
    max_ym = cursor.fetchone()[0]
    
    lots = {"양상추": [], "양배추": [], "적채": [], "프릴아이스": []}
    if max_ym:
        cursor.execute(
            "SELECT item_name, yymmdd, avail_g, vendor FROM closed_lots WHERE closed_ym = ?",
            (max_ym,)
        )
        rows = cursor.fetchall()
        for item_name, yymmdd, avail_g, vendor in rows:
            if item_name in lots:
                lots[item_name].append({
                    "yymmdd": yymmdd,
                    "avail_g": avail_g,
                    "vendor": vendor,
                    "is_from_db": True
                })
    conn.close()
    return max_ym, lots

def save_closed_lots(closed_ym, item_lots):
    """당월 마감 시, 남은 기말 롯트 재고를 마감 DB에 확정 저장합니다."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM closed_lots WHERE closed_ym = ?", (closed_ym,))
    
    for item_name, lot_list in item_lots.items():
        for lot in lot_list:
            if lot["avail_g"] > 0:
                cursor.execute(
                    "INSERT INTO closed_lots (item_name, yymmdd, avail_g, vendor, closed_ym) VALUES (?, ?, ?, ?, ?)",
                    (item_name, lot["yymmdd"], lot["avail_g"], lot.get("vendor", "K1"), closed_ym)
                )
    conn.commit()
    conn.close()

def import_db_from_file(uploaded_db_file, closed_ym):
    """사용자가 직접 업로드한 전월 마감 재고 엑셀/CSV를 DB로 로드합니다."""
    try:
        if uploaded_db_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_db_file)
        else:
            df = pd.read_excel(uploaded_db_file)

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM closed_lots WHERE closed_ym = ?", (closed_ym,))

        for _, row in df.iterrows():
            row_str = " ".join(row.dropna().astype(str).tolist())
            item_name = "양상추" if "양상추" in row_str else ("양배추" if "양배추" in row_str else ("적채" if "적채" in row_str else "프릴아이스"))
            
            m_date = re.search(r"\b(\d{6})\b", row_str)
            yymmdd = m_date.group(1) if m_date else datetime.now().strftime("%y%m%d")
            
            nums = [float(val) for val in row if isinstance(val, (int, float)) and val > 0]
            avail_g = (nums[0] * 1000.0) if nums else 0.0

            m_v = re.search(r"\b([A-Z]\d{1,2})\b", row_str)
            vendor = m_v.group(1) if m_v else "K1"

            if avail_g > 0:
                cursor.execute(
                    "INSERT INTO closed_lots (item_name, yymmdd, avail_g, vendor, closed_ym) VALUES (?, ?, ?, ?, ?)",
                    (item_name, yymmdd, avail_g, vendor, closed_ym)
                )
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.sidebar.error(f"DB 파일 처리 중 오류 발생: {e}")
        return False

init_db()

# ---------------------------------------------------------
# UI 및 인쇄 스타일 설정
# ---------------------------------------------------------
st.markdown("""
    <style>
    @media print {
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"], .stAlert, .no-print, [data-testid="stToolbar"] {
            display: none !important;
        }
        [data-testid="stDataFrame"] { display: none !important; }
        html, body, .main, .block-container, div, span, td, th {
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table-container { display: block !important; width: 100% !important; }
        .print-table {
            width: 100%; border-collapse: collapse; font-size: 10px;
            background-color: #ffffff !important; color: #000000 !important;
        }
        .print-table th, .print-table td {
            border: 1px solid #333333 !important; padding: 5px 4px;
            text-align: center; white-space: nowrap;
        }
        .print-table th { background-color: #f2f2f2 !important; font-weight: bold; }
    }
    .print-table-container { display: none; }
    </style>
""", unsafe_allow_html=True)

st.title("🥗 샐러드 생산일보 원료코드 대조 및 수불부 연동 System")

# ---------------------------------------------------------
# 사이드바: 전월 마감 DB 관리
# ---------------------------------------------------------
with st.sidebar:
    st.header("🗄️ 마스터 수불 DB 관리")
    last_ym, db_lots = get_latest_closed_lots()
    
    if last_ym:
        st.success(f"📌 현재 로드된 마감월: **{last_ym}**")
        with st.expander("이월 잔여 롯트 상세"):
            for k, v in db_lots.items():
                if v:
                    st.caption(f"**{k}**")
                    for l in v:
                        st.write(f"- 입고일: {l['yymmdd']} | 잔여: {round(l['avail_g']/1000.0, 1)}kg")
    else:
        st.warning("⚠️ 등록된 전월 마감 DB가 없습니다.\n아래에서 전월 마감 엑셀/CSV를 업로드해 주세요.")

    st.markdown("---")
    st.subheader("📥 3. 전월 마감 재고 DB 업로드 (수동 등록/초기화)")
    init_ym_input = st.text_input("업로드할 마감월 지정 (YYYY-MM)", value="2024-02")
    uploaded_init_db = st.file_uploader("전월 이월 재고 롯트 파일 (.csv, .xlsx)", type=["csv", "xlsx", "xls"], key="init_db_uploader")
    
    if uploaded_init_db is not None and st.button("전월 마감 DB 수동 등록", type="secondary"):
        if import_db_from_file(uploaded_init_db, init_ym_input):
            st.success(f"✅ {init_ym_input} 전월 마감 DB가 등록되었습니다!")
            st.rerun()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. 샐러드 생산일지 (Excel) 업로드 (당월 재고/매입자료 포함)", type=["xlsx", "xls"])
with col2:
    csv_file = st.file_uploader("2. 원료 매입 입고자료 / 수불부 (CSV/Excel) 업로드", type=["csv", "xlsx", "xls"])

prefix_map = {"양상추": "X", "양배추": "D", "적채": "A", "프릴": "F", "프릴아이스": "F"}

def extract_yymmdd_and_mmdd(val):
    if pd.isna(val):
        today = datetime.now()
        return today.strftime("%y"), today.strftime("%y%m%d"), today.strftime("%m%d")
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%y"), val.strftime("%y%m%d"), val.strftime("%m%d")

    val_str = str(val).strip()
    m_full = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        yy = m_full.group(1)[2:]
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    try:
        dt = pd.to_datetime(val_str, errors="coerce", yearfirst=True)
        if pd.notna(dt):
            return dt.strftime("%y"), dt.strftime("%y%m%d"), dt.strftime("%m%d")
    except Exception:
        pass

    current_yy = datetime.now().strftime("%y")
    m_short = re.search(r"(\d{1,2})[-./](\d{1,2})", val_str)
    if m_short:
        p1, p2 = m_short.group(1).zfill(2), m_short.group(2).zfill(2)
        yy, mm, dd = (p1, p2, "01") if int(p1) > 12 else (current_yy, p1, p2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]

    today = datetime.now()
    return current_yy, today.strftime("%y%m%d"), today.strftime("%m%d")

def find_stock_sheet_name(xls):
    for sheet_name in xls.sheet_names:
        if sheet_name.strip().lower() == "sheet1":
            return sheet_name
    keywords = ["양상추", "양배추", "적채"]
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=10)
            for idx, row in df.iterrows():
                row_str = " ".join(row.dropna().astype(str).tolist())
                if all(k in row_str for k in keywords):
                    return sheet_name
        except Exception:
            continue
    return xls.sheet_names[0]

# ---------------------------------------------------------
# [완벽 수정] openpyxl로 A~J열 (0~9 인덱스) 데이터 직접 파싱
# A(0):날짜 | B(1):양상추매입 | C(2):양상추기말 | D(3):양상추투입
# E(4):양배추매입 | F(5):양배추기말 | G(6):양배추투입
# H(7):적채매입   | I(8):적채기말   | J(9):적채투입
# ---------------------------------------------------------
def parse_and_convert_to_standard_ledger(excel_file):
    wb_in = openpyxl.load_workbook(excel_file, data_only=True)
    stock_sheet = find_stock_sheet_name(pd.ExcelFile(excel_file))
    ws_in = wb_in[stock_sheet]

    def parse_cell(val):
        if val is None: return 0.0
        if isinstance(val, (int, float)): return float(val)
        try:
            c = re.sub(r"[^\d.]", "", str(val))
            return float(c) if c else 0.0
        except Exception:
            return 0.0

    raw_parsed_list = []

    for row in ws_in.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None: continue
        raw_date = row[0]
        if "날짜" in str(raw_date): continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(raw_date)
        formatted_date = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"

        y_in  = parse_cell(row[1]) if len(row) > 1 else 0.0
        y_stk = parse_cell(row[2]) if len(row) > 2 else 0.0
        y_use = parse_cell(row[3]) if len(row) > 3 else 0.0

        b_in  = parse_cell(row[4]) if len(row) > 4 else 0.0
        b_stk = parse_cell(row[5]) if len(row) > 5 else 0.0
        b_use = parse_cell(row[6]) if len(row) > 6 else 0.0

        j_in  = parse_cell(row[7]) if len(row) > 7 else 0.0
        j_stk = parse_cell(row[8]) if len(row) > 8 else 0.0
        j_use = parse_cell(row[9]) if len(row) > 9 else 0.0

        raw_parsed_list.append({
            "date": formatted_date, "yymmdd": yymmdd,
            "y_in": y_in, "y_use": y_use, "y_stk": y_stk,
            "b_in": b_in, "b_use": b_use, "b_stk": b_stk,
            "j_in": j_in, "j_use": j_use, "j_stk": j_stk
        })

    # 동일 날짜 행 통합 (합산)
    grouped_by_date = {}
    date_order = []

    for item in raw_parsed_list:
        dt = item["date"]
        if dt not in grouped_by_date:
            grouped_by_date[dt] = {
                "date": dt, "yymmdd": item["yymmdd"],
                "y_in": 0.0, "y_use": 0.0, "y_stk": 0.0,
                "b_in": 0.0, "b_use": 0.0, "b_stk": 0.0,
                "j_in": 0.0, "j_use": 0.0, "j_stk": 0.0
            }
            date_order.append(dt)

        grouped_by_date[dt]["y_in"] += item["y_in"]
        grouped_by_date[dt]["b_in"] += item["b_in"]
        grouped_by_date[dt]["j_in"] += item["j_in"]

        grouped_by_date[dt]["y_use"] += item["y_use"]
        grouped_by_date[dt]["b_use"] += item["b_use"]
        grouped_by_date[dt]["j_use"] += item["j_use"]

        if item["y_stk"] > 0: grouped_by_date[dt]["y_stk"] = item["y_stk"]
        if item["b_stk"] > 0: grouped_by_date[dt]["b_stk"] = item["b_stk"]
        if item["j_stk"] > 0: grouped_by_date[dt]["j_stk"] = item["j_stk"]

    # FIFO 소진 큐 생성 (DB 매입 연동)
    _, db_lots = get_latest_closed_lots()
    item_queues = {
        "yangsangchu": copy.deepcopy(db_lots["양상추"]),
        "yangbaechu": copy.deepcopy(db_lots["양배추"]),
        "jeokchae": copy.deepcopy(db_lots["적채"])
    }

    data_rows = []
    generated_code_rows = []
    target_items_order = ["양상추", "양배추", "적채"]

    for dt in date_order:
        g = grouped_by_date[dt]
        yymmdd = g["yymmdd"]
        mmdd_str = f"{yymmdd[2:4]}/{yymmdd[4:]}"

        if g["y_in"] > 0: item_queues["yangsangchu"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["y_in"], "vendor": "K1"})
        if g["b_in"] > 0: item_queues["yangbaechu"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["b_in"], "vendor": "K1"})
        if g["j_in"] > 0: item_queues["jeokchae"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["j_in"], "vendor": "K1"})

        def trace_usage(item_key, item_prefix, use_amount):
            if use_amount <= 0: return "-", "-"
            rem = use_amount
            used_sources, used_codes = [], []

            for lot in item_queues[item_key]:
                if rem <= 0: break
                if lot["avail"] <= 0: continue
                u = min(lot["avail"], rem)
                lot["avail"] -= u
                rem -= u

                vendor = lot.get("vendor", "K1")
                used_sources.append(f"{lot['label']}({round(u/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{lot['yymmdd']}-{vendor}")

            if rem > 0:
                used_sources.append(f"부족({round(rem/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{yymmdd}-K1[부족]")

            return " / ".join(used_sources), " / ".join(list(dict.fromkeys(used_codes)))

        y_src, y_code = trace_usage("yangsangchu", "X", g["y_use"])
        b_src, b_code = trace_usage("yangbaechu", "D", g["b_use"])
        j_src, j_code = trace_usage("jeokchae", "A", g["j_use"])

        note = "휴무" if (g["y_in"] == 0 and g["y_use"] == 0 and g["b_in"] == 0 and g["b_use"] == 0 and g["j_in"] == 0 and g["j_use"] == 0) else ""

        row_data = {
            "date": dt, "yymmdd": yymmdd,
            "yangsangchu_in": g["y_in"], "yangsangchu_use": g["y_use"], "yangsangchu_source": y_src, "yangsangchu_stock": g["y_stk"],
            "yangbaechu_in": g["b_in"], "yangbaechu_use": g["b_use"], "yangbaechu_source": b_src, "yangbaechu_stock": g["b_stk"],
            "jeokchae_in": g["j_in"], "jeokchae_use": g["j_use"], "jeokchae_source": j_src, "jeokchae_stock": g["j_stk"],
            "note": note
        }
        data_rows.append(row_data)

        code_map = {"양상추": y_code, "양배추": b_code, "적채": j_code}
        src_map = {"양상추": y_src, "양배추": b_src, "적채": j_src}
        use_map = {"양상추": g["y_use"], "양배추": g["b_use"], "적채": g["j_use"]}

        for item_name in target_items_order:
            u_val = use_map[item_name]
            if u_val > 0:
                generated_code_rows.append({
                    "날짜": dt, "원재료": item_name, "생성코드": code_map.get(item_name, "-"),
                    "투입량(kg)": round(u_val / 1000.0, 1), "준비량(kg)": 0.0, "소진매입일(상세)": src_map.get(item_name, "-"), "비고": note
                })

    # --- openpyxl 엑셀 생성 ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "원재료 수불부"
    ws.views.sheetView[0].showGridLines = True
    ws.merge_cells("A1:Q1")
    ws["A1"] = "원재료 수불부 (Raw Material Inventory Ledger - FIFO 롯트 추적)"
    ws["A1"].font = Font(name="맑은 고딕", size=16, bold=True, color="1F4E79")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    header_fill_main = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    regular_font = Font(name="맑은 고딕", size=10)
