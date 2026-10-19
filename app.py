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

init_db()

# ---------------------------------------------------------
# UI 및 스타일 설정
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
# [수정] 엑셀 행 순서 파싱 및 날짜별 병합 누적 로직 완벽 보완
# ---------------------------------------------------------
def parse_and_convert_to_standard_ledger(excel_file):
    xls = pd.ExcelFile(excel_file)
    stock_sheet = find_stock_sheet_name(xls)
    df_raw = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

    def parse_num(val):
        if pd.isna(val): return 0.0
        try:
            c = re.sub(r"[^\d.]", "", str(val))
            return float(c) if c else 0.0
        except Exception: return 0.0

    # 1단계: 순서 보장 날짜별 데이터 수집
    raw_parsed_list = []

    for idx, row in df_raw.iterrows():
        if len(row) == 0: continue
        raw_date = row.iloc[0]
        
        if pd.isna(raw_date) or str(raw_date).strip() == "" or "날짜" in str(raw_date):
            continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(raw_date)
        formatted_date = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"

        # B:매입(1), C:기말재고(2), D:투입(3)
        y_in  = parse_num(row.iloc[1]) if len(row) > 1 else 0.0
        y_stk = parse_num(row.iloc[2]) if len(row) > 2 else 0.0
        y_use = parse_num(row.iloc[3]) if len(row) > 3 else 0.0

        # E:매입(4), F:기말재고(5), G:투입(6)
        b_in  = parse_num(row.iloc[4]) if len(row) > 4 else 0.0
        b_stk = parse_num(row.iloc[5]) if len(row) > 5 else 0.0
        b_use = parse_num(row.iloc[6]) if len(row) > 6 else 0.0

        # H:매입(7), I:기말재고(8), J:투입(9)
        j_in  = parse_num(row.iloc[7]) if len(row) > 7 else 0.0
        j_stk = parse_num(row.iloc[8]) if len(row) > 8 else 0.0
        j_use = parse_num(row.iloc[9]) if len(row) > 9 else 0.0

        raw_parsed_list.append({
            "date": formatted_date, "yymmdd": yymmdd,
            "y_in": y_in, "y_use": y_use, "y_stk": y_stk,
            "b_in": b_in, "b_use": b_use, "b_stk": b_stk,
            "j_in": j_in, "j_use": j_use, "j_stk": j_stk
        })

    # 2단계: 동일 날짜 행 통합 (합산 처리)
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

        # 기말재고가 있는 행의 수량 보존
        if item["y_stk"] > 0: grouped_by_date[dt]["y_stk"] = item["y_stk"]
        if item["b_stk"] > 0: grouped_by_date[dt]["b_stk"] = item["b_stk"]
        if item["j_stk"] > 0: grouped_by_date[dt]["j_stk"] = item["j_stk"]

    # 3단계: FIFO 소진 큐 생성 (DB 매입 우선)
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

        # 당일 입고 롯트 FIFO 큐에 등록
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
    thin_border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"),
                         top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    headers_main = [("A2:A3", "일자"), ("B2:F2", "양상추"), ("G2:K2", "양배추"), ("L2:P2", "적채"), ("Q2:Q3", "비고")]
    for cell_range, title in headers_main:
        ws.merge_cells(cell_range)
        ws[cell_range.split(":")[0]] = title

    for col_idx in range(1, 18):
        c = ws.cell(row=2, column=col_idx)
        c.fill, c.font, c.alignment = header_fill_main, Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF"), Alignment(horizontal="center", vertical="center")

    sub_headers = ["기초재고", "입고량", "투입량", "소진 매입일자(상세)", "기말재고"]
    for start_col in [2, 7, 12]:
        for idx, sub in enumerate(sub_headers):
            c = ws.cell(row=3, column=start_col + idx, value=sub)
            c.fill, c.font, c.alignment = header_fill_sub, Font(name="맑은 고딕", size=10, bold=True, color="1F4E79"), Alignment(horizontal="center", vertical="center")

    start_row = 4
    for r_idx, item in enumerate(data_rows, start=start_row):
        ws.cell(row=r_idx, column=1, value=item["date"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=17, value=item.get("note", "")).alignment = Alignment(horizontal="left", vertical="center")

        items_map = [
            (2, item.get("yangsangchu_in", 0), item.get("yangsangchu_use", 0), item.get("yangsangchu_source", "-"), item.get("yangsangchu_stock", 0)),
            (7, item.get("yangbaechu_in", 0), item.get("yangbaechu_use", 0), item.get("yangbaechu_source", "-"), item.get("yangbaechu_stock", 0)),
            (12, item.get("jeokchae_in", 0), item.get("jeokchae_use", 0), item.get("jeokchae_source", "-"), item.get("jeokchae_stock", 0))
        ]

        for start_col, in_val, use_val, source_str, stock_val in items_map:
            end_col_let = get_column_letter(start_col + 4)
            if r_idx == start_row:
                ws.cell(row=r_idx, column=start_col, value=stock_val + use_val - in_val)
            else:
                ws.cell(row=r_idx, column=start_col, value=f"={end_col_let}{r_idx-1}")

            ws.cell(row=r_idx, column=start_col + 1, value=in_val)
            ws.cell(row=r_idx, column=start_col + 2, value=use_val)
            ws.cell(row=r_idx, column=start_col + 3, value=source_str).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(row=r_idx, column=start_col + 4, value=stock_val)

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
# Streamlit 화면 구성
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(["📊 표준 원재료 수불부 정리", "🏷️ 원료코드 생성 목록", "🔒 당월 수불 마감 확정"])

if excel_file is not None:
    try:
        excel_bytes, standard_rows, code_summary_rows, stock_sheet_name, final_lots = parse_and_convert_to_standard_ledger(excel_file)

        with tab1:
            st.subheader("📋 표준 원재료 수불부 (수불부 투입량 정확 반영)")
            df_preview = pd.DataFrame(standard_rows)
            if "yymmdd" in df_preview.columns: df_preview = df_preview.drop(columns=["yymmdd"])

            ordered_cols = [
                "date",
                "yangsangchu_in", "yangsangchu_use", "yangsangchu_source", "yangsangchu_stock",
                "yangbaechu_in", "yangbaechu_use", "yangbaechu_source", "yangbaechu_stock",
                "jeokchae_in", "jeokchae_use", "jeokchae_source", "jeokchae_stock",
                "note"
            ]
            df_preview = df_preview[[col for col in ordered_cols if col in df_preview.columns]]

            korean_column_map = {
                "date": "일자",
                "yangsangchu_in": "양상추_입고량", "yangsangchu_use": "양상추_투입량", "yangsangchu_source": "양상추_소진 매입일자(상세)", "yangsangchu_stock": "양상추_기말재고",
                "yangbaechu_in": "양배추_입고량", "yangbaechu_use": "양배추_투입량", "yangbaechu_source": "양배추_소진 매입일자(상세)", "yangbaechu_stock": "양배추_기말재고",
                "jeokchae_in": "적채_입고량", "jeokchae_use": "적채_투입량", "jeokchae_source": "적채_소진 매입일자(상세)", "jeokchae_stock": "적채_기말재고",
                "note": "비고",
            }
            df_preview = df_preview.rename(columns=korean_column_map)

            st.dataframe(df_preview, use_container_width=True)
            st.download_button("📥 정제된 표준 원재료 수불부 엑셀 다운로드 (.xlsx)", data=excel_bytes, file_name="표준_원재료_수불부.xlsx", use_container_width=True)

        with tab2:
            st.subheader("🏷️ 소진 매입재고 기반 일자별 원료코드 목록")
            df_code = pd.DataFrame(code_summary_rows)
            st.dataframe(df_code, use_container_width=True)

        with tab3:
            st.subheader("🔒 당월 수불 마감 및 전월 DB 확정")
            last_date_str = standard_rows[-1]["date"] if standard_rows else datetime.now().strftime("%Y-%m-%d")
            closed_ym_input = st.text_input("마감 확정 월 (YYYY-MM)", value=last_date_str[:7])

            if st.button("🔒 당월 수불 마감 및 DB 확정 저장", type="primary"):
                save_closed_lots(closed_ym_input, final_lots)
                st.success(f"✅ **{closed_ym_input}** 수불 마감이 저장되었습니다!")

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    st.info("👆 상단에서 1. 샐러드 생산일지 엑셀 파일을 업로드해 주세요.")
