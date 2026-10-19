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

        # 표준 컬럼 매핑 추정
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM closed_lots WHERE closed_ym = ?", (closed_ym,))

        for _, row in df.iterrows():
            row_str = " ".join(row.dropna().astype(str).tolist())
            item_name = "양상추" if "양상추" in row_str else ("양배추" if "양배추" in row_str else ("적채" if "적채" in row_str else "프릴아이스"))
            
            # 날짜(yymmdd) 추출
            m_date = re.search(r"\b(\d{6})\b", row_str)
            yymmdd = m_date.group(1) if m_date else datetime.now().strftime("%y%m%d")
            
            # 수량(kg/g) 추출 - 숫자형 값 탐색
            nums = [float(val) for val in row if isinstance(val, (int, float)) and val > 0]
            avail_g = (nums[0] * 1000.0) if nums else 0.0

            # 공급사 코드
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
# 사이드바: 전월 마감 DB 업로드 및 관리 기능
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
        st.warning("⚠️ 등록된 전월 마감 DB가 없습니다.\n아래에서 최초 전월 마감 엑셀/CSV를 업로드해 주세요.")

    st.markdown("---")
    st.subheader("📥 3. 전월 마감 재고 DB 업로드 (최초 1회 또는 초기화)")
    init_ym_input = st.text_input("업로드할 마감월 지정 (YYYY-MM)", value="2024-02")
    uploaded_init_db = st.file_uploader("전월 이월 재고 롯트 파일 (.csv, .xlsx)", type=["csv", "xlsx", "xls"], key="init_db_uploader")
    
    if uploaded_init_db is not None and st.button("전월 마감 DB 수동 등록", type="secondary"):
        if import_db_from_file(uploaded_init_db, init_ym_input):
            st.success(f"✅ {init_ym_input} 전월 마감 DB가 등록되었습니다!")
            st.rerun()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. 샐러드 생산일지 (Excel) 업로드", type=["xlsx", "xls"])
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
# 수불부 정제 및 DB 기반 FIFO 소진 추적
# ---------------------------------------------------------
def parse_and_convert_to_standard_ledger(excel_file):
    xls = pd.ExcelFile(excel_file)
    stock_sheet = find_stock_sheet_name(xls)
    df_raw = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

    def parse_cell(val):
        if pd.isna(val): return 0.0
        try:
            c = re.sub(r"[^\d.]", "", str(val))
            return float(c) if c else 0.0
        except Exception: return 0.0

    pending_in = {"yangsangchu": 0.0, "yangbaechu": 0.0, "jeokchae": 0.0}
    raw_data_rows = []
    is_first_row = True

    for idx, row in df_raw.iterrows():
        if len(row) == 0 or pd.isna(row.iloc[0]): continue
        a_val = str(row.iloc[0]).strip()
        if not a_val or a_val.lower() == "nan" or "날짜" in a_val or "매입" in a_val: continue

        if "입고" in a_val:
            pending_in["yangsangchu"] += parse_cell(row.iloc[1]) if len(row) > 1 else 0.0
            pending_in["yangbaechu"] += parse_cell(row.iloc[3]) if len(row) > 3 else 0.0
            pending_in["jeokchae"] += parse_cell(row.iloc[5]) if len(row) > 5 else 0.0
            continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(row.iloc[0])
        formatted_date = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"

        y_stock = parse_cell(row.iloc[1]) if len(row) > 1 else 0.0
        y_use = parse_cell(row.iloc[2]) if len(row) > 2 else 0.0
        b_stock = parse_cell(row.iloc[3]) if len(row) > 3 else 0.0
        b_use = parse_cell(row.iloc[4]) if len(row) > 4 else 0.0
        j_stock = parse_cell(row.iloc[5]) if len(row) > 5 else 0.0
        j_use = parse_cell(row.iloc[6]) if len(row) > 6 else 0.0

        y_in, b_in, j_in = pending_in["yangsangchu"], pending_in["yangbaechu"], pending_in["jeokchae"]
        pending_in = {"yangsangchu": 0.0, "yangbaechu": 0.0, "jeokchae": 0.0}

        if is_first_row:
            y_beg, b_beg, j_beg = y_stock, b_stock, j_stock
            is_first_row = False
        else:
            y_beg = b_beg = j_beg = 0.0

        note = "휴무" if (y_in == 0 and y_use == 0 and b_in == 0 and b_use == 0 and j_in == 0 and j_use == 0 and y_stock == 0 and b_stock == 0 and j_stock == 0) else ""

        raw_data_rows.append({
            "date": formatted_date, "yymmdd": yymmdd,
            "yangsangchu_beg": y_beg, "yangsangchu_in": y_in, "yangsangchu_use": y_use, "yangsangchu_stock": y_stock,
            "yangbaechu_beg": b_beg, "yangbaechu_in": b_in, "yangbaechu_use": b_use, "yangbaechu_stock": b_stock,
            "jeokchae_beg": j_beg, "jeokchae_in": j_in, "jeokchae_use": j_use, "jeokchae_stock": j_stock,
            "note": note
        })

    # --- FIFO 큐 구성 (전월 마감 DB 연결) ---
    _, db_lots = get_latest_closed_lots()
    item_queues = {
        "yangsangchu": copy.deepcopy(db_lots["양상추"]),
        "yangbaechu": copy.deepcopy(db_lots["양배추"]),
        "jeokchae": copy.deepcopy(db_lots["적채"])
    }

    # DB에 잔여 롯트가 없을 경우 파일 내 기초재고 활용
    if raw_data_rows:
        f_row = raw_data_rows[0]
        if not item_queues["yangsangchu"] and f_row["yangsangchu_beg"] > 0:
            item_queues["yangsangchu"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["yangsangchu_beg"], "vendor": "K1"})
        if not item_queues["yangbaechu"] and f_row["yangbaechu_beg"] > 0:
            item_queues["yangbaechu"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["yangbaechu_beg"], "vendor": "K1"})
        if not item_queues["jeokchae"] and f_row["jeokchae_beg"] > 0:
            item_queues["jeokchae"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["jeokchae_beg"], "vendor": "K1"})

    for k in item_queues:
        for q in item_queues[k]:
            if "avail" not in q:
                q["avail"] = q.get("avail_g", 0.0)
            if "label" not in q:
                q["label"] = f"{q['yymmdd'][2:4]}/{q['yymmdd'][4:]} 이월"

    data_rows = []
    generated_code_rows = []
    target_items_order = ["양상추", "양배추", "적채"]

    for row in raw_data_rows:
        row_copy = copy.deepcopy(row)
        mmdd_str = f"{row['yymmdd'][2:4]}/{row['yymmdd'][4:]}"

        if row["yangsangchu_in"] > 0:
            item_queues["yangsangchu"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["yangsangchu_in"], "vendor": "K1"})
        if row["yangbaechu_in"] > 0:
            item_queues["yangbaechu"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["yangbaechu_in"], "vendor": "K1"})
        if row["jeokchae_in"] > 0:
            item_queues["jeokchae"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["jeokchae_in"], "vendor": "K1"})

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
                used_codes.append(f"{item_prefix}{row['yymmdd']}-K1[부족]")

            return " / ".join(used_sources), " / ".join(list(dict.fromkeys(used_codes)))

        y_src, y_code = trace_usage("yangsangchu", "X", row["yangsangchu_use"])
        b_src, b_code = trace_usage("yangbaechu", "D", row["yangbaechu_use"])
        j_src, j_code = trace_usage("jeokchae", "A", row["jeokchae_use"])

        row_copy["yangsangchu_source"], row_copy["yangbaechu_source"], row_copy["jeokchae_source"] = y_src, b_src, j_src
        data_rows.append(row_copy)

        code_map = {"양상추": y_code, "양배추": b_code, "적채": j_code}
        src_map = {"양상추": y_src, "양배추": b_src, "적채": j_src}

        for item_name in target_items_order:
            use_val = row["yangsangchu_use"] if item_name == "양상추" else (row["yangbaechu_use"] if item_name == "양배추" else row["jeokchae_use"])
            if use_val > 0:
                generated_code_rows.append({
                    "날짜": row["date"], "원재료": item_name, "생성코드": code_map.get(item_name, "-"),
                    "투입량(kg)": round(use_val / 1000.0, 1), "소진매입일(상세)": src_map.get(item_name, "-"), "비고": row["note"]
                })

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
    header_font_main = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    header_font_sub = Font(name="맑은 고딕", size=10, bold=True, color="1F4E79")
    regular_font = Font(name="맑은 고딕", size=10)
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    headers_main = [("A2:A3", "일자"), ("B2:F2", "양상추"), ("G2:K2", "양배추"), ("L2:P2", "적채"), ("Q2:Q3", "비고")]
    for cell_range, title in headers_main:
        ws.merge_cells(cell_range)
        ws[cell_range.split(":")[0]] = title

    for col_idx in range(1, 18):
        c = ws.cell(row=2, column=col_idx)
        c.fill, c.font, c.alignment = header_fill_main, header_font_main, Alignment(horizontal="center", vertical="center")

    sub_headers = ["기초재고", "입고량", "투입량", "소진 매입일자(상세)", "기말재고"]
    for start_col in [2, 7, 12]:
        for idx, sub in enumerate(sub_headers):
            c = ws.cell(row=3, column=start_col + idx, value=sub)
            c.fill, c.font, c.alignment = header_fill_sub, header_font_sub, Alignment(horizontal="center", vertical="center")

    start_row = 4
    for r_idx, item in enumerate(data_rows, start=start_row):
        ws.cell(row=r_idx, column=1, value=item["date"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=17, value=item.get("note", "")).alignment = Alignment(horizontal="left", vertical="center")

        items_map = [
            (2, item.get("yangsangchu_beg", 0), item.get("yangsangchu_in", 0), item.get("yangsangchu_use", 0), item.get("yangsangchu_source", "-")),
            (7, item.get("yangbaechu_beg", 0), item.get("yangbaechu_in", 0), item.get("yangbaechu_use", 0), item.get("yangbaechu_source", "-")),
            (12, item.get("jeokchae_beg", 0), item.get("jeokchae_in", 0), item.get("jeokchae_use", 0), item.get("jeokchae_source", "-"))
        ]

        for start_col, beg_val, in_val, use_val, source_str in items_map:
            end_col_let = get_column_letter(start_col + 4)
            if r_idx == start_row: ws.cell(row=r_idx, column=start_col, value=beg_val)
            else: ws.cell(row=r_idx, column=start_col, value=f"={end_col_let}{r_idx-1}")

            ws.cell(row=r_idx, column=start_col + 1, value=in_val)
            ws.cell(row=r_idx, column=start_col + 2, value=use_val)
            ws.cell(row=r_idx, column=start_col + 3, value=source_str).alignment = Alignment(horizontal="center", vertical="center")

            c_beg, c_in, c_use = get_column_letter(start_col), get_column_letter(start_col + 1), get_column_letter(start_col + 2)
            ws.cell(row=r_idx, column=start_col + 4, value=f"={c_beg}{r_idx}+{c_in}{r_idx}-{c_use}{r_idx}")

        for c_idx in range(1, 18):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font, cell.border = regular_font, thin_border
            if c_idx in [2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")

    buffer = io.BytesIO()
    wb.save(buffer)

    final_remaining_lots = {
        "양상추": [{"yymmdd": q["yymmdd"], "avail_g": q["avail"], "vendor": q.get("vendor", "K1")} for q in item_queues["yangsangchu"] if q["avail"] > 0],
        "양배추": [{"yymmdd": q["yymmdd"], "avail_g": q["avail"], "vendor": q.get("vendor", "K1")} for q in item_queues["yangbaechu"] if q["avail"] > 0],
        "적채": [{"yymmdd": q["yymmdd"], "avail_g": q["avail"], "vendor": q.get("vendor", "K1")} for q in item_queues["jeokchae"] if q["avail"] > 0]
    }

    return buffer.getvalue(), data_rows, generated_code_rows, stock_sheet, final_remaining_lots

# ---------------------------------------------------------
# Streamlit Tab UI
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 표준 원재료 수불부", "🏷️ 일자별 원료코드 생성", "🔒 당월 수불 마감 확정"])

if excel_file is not None:
    try:
        excel_bytes, standard_rows, code_summary_rows, stock_sheet_name, final_lots = parse_and_convert_to_standard_ledger(excel_file)

        with tab1:
            st.subheader("📋 표준 원재료 수불부 (전월 마감 롯트 연동)")
            df_preview = pd.DataFrame(standard_rows)
            if "yymmdd" in df_preview.columns: df_preview = df_preview.drop(columns=["yymmdd"])
            st.dataframe(df_preview, use_container_width=True)
            st.download_button("📥 수불부 엑셀 다운로드 (.xlsx)", data=excel_bytes, file_name="표준_원재료_수불부.xlsx", use_container_width=True)

        with tab2:
            st.subheader("🏷️ 소진 매입재고 기반 일자별 원료코드 목록")
            df_code = pd.DataFrame(code_summary_rows)
            st.dataframe(df_code, use_container_width=True)

        with tab3:
            st.subheader("🔒 당월 수불 마감 및 전월 DB 확정")
            st.info("당월 수불 검토가 완료되면 아래 마감 버튼을 눌러 **기말 잔여 롯트를 DB에 저장/확정**하세요.\n다음 달 수불 계산 시 이 데이터가 '기초 이월 롯트'로 자동 연동됩니다.")
            
            last_date_str = standard_rows[-1]["date"] if standard_rows else datetime.now().strftime("%Y-%m-%d")
            default_ym = last_date_str[:7]
            closed_ym_input = st.text_input("마감 확정 월 (YYYY-MM)", value=default_ym)

            st.write("▼ **마감 후 다음 달로 이월될 잔여 롯트 미리보기**")
            for item_k, l_list in final_lots.items():
                st.write(f"• **{item_k}**: {len(l_list)}개 롯트 잔여")
                for l in l_list:
                    st.caption(f"   - 입고일: {l['yymmdd']} | 잔여량: {round(l['avail_g']/1000.0, 1)}kg")

            if st.button("🔒 당월 수불 마감 및 DB 확정 저장", type="primary"):
                save_closed_lots(closed_ym_input, final_lots)
                st.balloons()
                st.success(f"✅ **{closed_ym_input}** 수불 마감이 성공적으로 저장되었습니다! 다음 달 수불 계산 시 반영됩니다.")

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    st.info("👆 상단에서 샐러드 생산일지 엑셀 파일을 업로드해 주세요.")
