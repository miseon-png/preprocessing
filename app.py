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
import streamlit.components.v1 as components

st.set_page_config(page_title="샐러드 생산일보 및 원료코드 System", layout="wide")

DB_FILE = "stock_master.db"

# ---------------------------------------------------------
# DB 관리 함수
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

def clear_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM closed_lots")
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
                    "avail": avail_g,
                    "vendor": vendor,
                    "label": f"{yymmdd[2:4]}/{yymmdd[4:]} 이월",
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
            avail_val = lot.get("avail", lot.get("avail_g", 0.0))
            if avail_val > 0:
                cursor.execute(
                    "INSERT INTO closed_lots (item_name, yymmdd, avail_g, vendor, closed_ym) VALUES (?, ?, ?, ?, ?)",
                    (item_name, lot["yymmdd"], avail_val, lot.get("vendor", "K01"), closed_ym)
                )
    conn.commit()
    conn.close()

init_db()

# ---------------------------------------------------------
# 프린터 전용 직접 인쇄 HTML 생성 함수
# ---------------------------------------------------------
def render_printable_document(title, df_data):
    html_table = df_data.to_html(index=False, classes="report-table")
    
    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: '맑은 고딕', sans-serif; padding: 20px; color: #000; background-color: #fff; }}
        .header-container {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; }}
        .report-title {{ font-size: 20pt; font-weight: bold; color: #1F4E79; margin: 0; }}
        
        .approval-table {{ border-collapse: collapse; text-align: center; font-size: 9pt; }}
        .approval-table th, .approval-table td {{ border: 1px solid #333; padding: 4px 8px; }}
        .approval-table th {{ background-color: #f2f2f2; width: 50px; }}
        .approval-table td {{ height: 40px; width: 60px; }}
        
        .report-table {{ width: 100%; border-collapse: collapse; font-size: 9pt; margin-top: 10px; }}
        .report-table th, .report-table td {{ border: 1px solid #333; padding: 6px 4px; text-align: center; }}
        .report-table th {{ background-color: #D9E1F2; font-weight: bold; color: #1F4E79; }}
        .report-table tr:nth-child(even) {{ background-color: #F9FAFB; }}
        
        .print-btn {{
            background-color: #1F4E79; color: white; padding: 8px 16px; border: none;
            border-radius: 4px; font-weight: bold; cursor: pointer; margin-bottom: 15px;
        }}
        @media print {{
            .print-btn {{ display: none !important; }}
            body {{ padding: 0; }}
        }}
    </style>
    </head>
    <body>
        <button class="print-btn" onclick="window.print()">🖨️ 이 문서 프린터로 바로 출력하기</button>
        <div class="header-container">
            <h1 class="report-title">{title}</h1>
            <table class="approval-table">
                <tr><th rowspan="2">결<br>재</th><th>담당</th><th>검토</th><th>승인</th></tr>
                <tr><td></td><td></td><td></td></tr>
            </table>
        </div>
        {html_table}
    </body>
    </html>
    """
    return full_html

# ---------------------------------------------------------
# [신규] 전처리 공정일지 전용 HTML 양식 렌더링 함수
# ---------------------------------------------------------
def render_preprocess_log_document(target_date, product_name, items_data):
    dt_obj = datetime.strptime(target_date, "%Y-%m-%d")
    year_str, month_str, day_str = dt_obj.strftime("%Y"), dt_obj.strftime("%m"), dt_obj.strftime("%d")

    rows_html = ""
    total_rows_count = max(12, len(items_data))

    for idx in range(total_rows_count):
        if idx < len(items_data):
            item_info = items_data[idx]
            p_name = product_name if idx == 0 else ""
            rowspan_attr = f' rowspan="{len(items_data)}"' if idx == 0 else ""
            
            p_cell = f'<td{rowspan_attr} style="font-weight:bold; background-color:#FAFAFA;">{product_name}</td>' if idx == 0 else ""
            if idx >= len(items_data) and idx == len(items_data):
                p_cell = '<td></td>'

            rows_html += f"""
            <tr>
                {p_cell}
                <td>{item_info['원료명']}</td>
                <td style="font-family: monospace; font-size:10pt;">{item_info['원료코드']}</td>
                <td>적 / 부</td>
                <td style="text-align:left; padding-left:8px;">{item_info['점검항목']}</td>
                <td>{item_info['절단규격']}</td>
                <td style="text-align:right; padding-right:8px;">{item_info['생산량(전)']} Kg</td>
                <td style="text-align:right; padding-right:8px;">{item_info['투입량(후)']} Kg</td>
                <td></td>
            </tr>
            """
        else:
            rows_html += """
            <tr>
                <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
            </tr>
            """

    full_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="utf-8">
    <style>
        body {{ font-family: '맑은 고딕', sans-serif; padding: 20px; color: #000; background-color: #fff; }}
        .header-box {{ display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }}
        .doc-title {{ font-size: 24pt; font-weight: bold; margin: 0; letter-spacing: 2px; }}
        
        .app-table {{ border-collapse: collapse; text-align: center; font-size: 9pt; width: 220px; }}
        .app-table th, .app-table td {{ border: 1px solid #000; padding: 3px; }}
        .app-table th {{ background-color: #f2f2f2; width: 60px; height: 20px; }}
        .app-table td {{ height: 45px; }}

        .meta-table {{ width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 9.5pt; }}
        .meta-table td {{ border: 1px solid #000; padding: 5px 10px; text-align: center; }}
        
        .info-table {{ width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 9pt; }}
        .info-table td {{ border: 1px solid #000; padding: 5px 8px; }}
        .info-bg {{ background-color: #f2f2f2; font-weight: bold; text-align: center; width: 140px; }}

        .main-table {{ width: 100%; border-collapse: collapse; font-size: 9pt; }}
        .main-table th, .main-table td {{ border: 1px solid #000; padding: 6px 4px; text-align: center; height: 26px; }}
        .main-table th {{ background-color: #f2f2f2; font-weight: bold; }}

        .print-btn {{
            background-color: #1F4E79; color: white; padding: 8px 16px; border: none;
            border-radius: 4px; font-weight: bold; cursor: pointer; margin-bottom: 15px;
        }}
        @media print {{
            .print-btn {{ display: none !important; }}
            body {{ padding: 0; }}
        }}
    </style>
    </head>
    <body>
        <button class="print-btn" onclick="window.print()">🖨️ 이 공정일지 프린터로 바로 출력하기</button>
        
        <div class="header-box">
            <h1 class="doc-title">전처리 공정일지</h1>
            <table class="app-table">
                <tr><th>작 성</th><th>검 토</th><th>승 인</th></tr>
                <tr><td></td><td></td><td></td></tr>
            </table>
        </div>

        <table class="meta-table">
            <tr>
                <td style="width:12%; background-color:#f2f2f2; font-weight:bold;">작성일자</td>
                <td style="width:58%; font-weight:bold; font-size:10.5pt;">{year_str} 년 &nbsp;&nbsp; {month_str} 월 &nbsp;&nbsp; {day_str} 일</td>
                <td style="width:10%; background-color:#f2f2f2; font-weight:bold;">작업시간</td>
                <td style="width:20%;"> : &nbsp;&nbsp; ~ &nbsp;&nbsp; : </td>
                <td style="width:8%; background-color:#f2f2f2; font-weight:bold;">작성자</td>
                <td style="width:12%;"></td>
            </tr>
        </table>

        <table class="info-table">
            <tr>
                <td class="info-bg" rowspan="2">원료코드<br>부여방법</td>
                <td style="text-align:center; font-weight:bold; width:220px;">원료 구분 + 입고일자 - 공급업체 구분</td>
                <td class="info-bg" style="width:110px;">원료 구분<br>코드</td>
                <td style="font-size:8.5pt;">카이피라(C), 프릴아이스(F), 버터헤드(B), 레드오크(O), 로메인(R), 케일(K), 적근대(S)<br>치커리(E), 적채(A), 양배추(D), 양상추(X)</td>
            </tr>
            <tr>
                <td style="background-color:#f2f2f2; font-weight:bold; text-align:center;">공급업체 구분</td>
                <td colspan="2" style="font-size:8.5pt;">구름(G01) , 에상스팜(H02), 승승장구(S01), 한스(K01)</td>
            </tr>
        </table>

        <table class="main-table">
            <thead>
                <tr>
                    <th rowspan="2" style="width:15%;">제품명</th>
                    <th rowspan="2" style="width:10%;">원료명</th>
                    <th colspan="7">작업원물 정보</th>
                </tr>
                <tr>
                    <th style="width:16%;">원료코드</th>
                    <th style="width:8%;">상태</th>
                    <th style="width:23%;">점검 항목</th>
                    <th style="width:11%;">절단규격</th>
                    <th style="width:9%;">생산량(전)</th>
                    <th style="width:9%;">투입량(후)</th>
                    <th style="width:6%;">비고</th>
                </tr>
            </thead>
            <tbody>
                {rows_html}
            </tbody>
        </table>
    </body>
    </html>
    """
    return full_html

st.title("🥗 샐러드 생산일보 원료코드 대조 및 수불부 연동 System")

# ---------------------------------------------------------
# 사이드바
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
        st.info("💡 등록된 전월 마감 DB가 없습니다.\n당월 업로드 엑셀의 입고분부터 순서대로 소진됩니다.")

    st.markdown("---")
    if st.button("🗑️ 과거 이월 DB 완전 초기화", type="secondary"):
        clear_db()
        st.success("✅ DB가 초기화되었습니다!")
        st.rerun()

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. 샐러드 생산일지 (Excel) 업로드", type=["xlsx", "xls"])
with col2:
    csv_file = st.file_uploader("2. 원료 매입 입고자료 / 수불부 (CSV/Excel) 업로드", type=["csv", "xlsx", "xls"])

prefix_map = {"양상추": "X", "양배추": "D", "적채": "A", "프릴": "F", "프릴아이스": "F"}
vendor_code_map = {"한스": "K01", "승승장구": "S01", "에상스팜": "H02", "넥스토팜": "N01"}

# ---------------------------------------------------------
# 날짜 정밀 정제 함수
# ---------------------------------------------------------
def extract_yymmdd_and_mmdd(val):
    if pd.isna(val):
        today = datetime.now()
        return today.strftime("%y"), today.strftime("%y%m%d"), today.strftime("%m%d")
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime("%y"), val.strftime("%y%m%d"), val.strftime("%m%d")

    val_str = str(val).strip()
    
    m_dot_digit = re.search(r"(\d{2})\.(\d{2})(\d{2})", val_str)
    if m_dot_digit:
        yy, mm, dd = m_dot_digit.group(1), m_dot_digit.group(2), m_dot_digit.group(3)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_full = re.search(r"(\d{2,4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        p1 = m_full.group(1)
        yy = p1[2:] if len(p1) == 4 else p1.zfill(2)
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]

    current_yy = datetime.now().strftime("%y")
    m_short = re.search(r"(\d{1,2})\s*[-./월]\s*(\d{1,2})", val_str)
    if m_short:
        mm = m_short.group(1).zfill(2)
        dd = m_short.group(2).zfill(2)
        return current_yy, f"{current_yy}{mm}{dd}", f"{mm}{dd}"

    try:
        dt = pd.to_datetime(val_str, errors="coerce", yearfirst=True)
        if pd.notna(dt):
            return dt.strftime("%y"), dt.strftime("%y%m%d"), dt.strftime("%m%d")
    except Exception:
        pass

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

    raw_parsed_list = []

    for idx, row in df_raw.iterrows():
        if len(row) == 0 or pd.isna(row.iloc[0]): continue
        a_val = str(row.iloc[0]).strip()
        if not a_val or "날짜" in a_val or "입고" in a_val or "매입" in a_val: continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(a_val)
        formatted_date = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"

        y_in  = parse_cell(row.iloc[1]) if len(row) > 1 else 0.0
        y_stk = parse_cell(row.iloc[2]) if len(row) > 2 else 0.0
        y_use = parse_cell(row.iloc[3]) if len(row) > 3 else 0.0

        b_in  = parse_cell(row.iloc[4]) if len(row) > 4 else 0.0
        b_stk = parse_cell(row.iloc[5]) if len(row) > 5 else 0.0
        b_use = parse_cell(row.iloc[6]) if len(row) > 6 else 0.0

        j_in  = parse_cell(row.iloc[7]) if len(row) > 7 else 0.0
        j_stk = parse_cell(row.iloc[8]) if len(row) > 8 else 0.0
        j_use = parse_cell(row.iloc[9]) if len(row) > 9 else 0.0

        raw_parsed_list.append({
            "date": formatted_date, "yymmdd": yymmdd,
            "y_in": y_in, "y_use": y_use, "y_stk": y_stk,
            "b_in": b_in, "b_use": b_use, "b_stk": b_stk,
            "j_in": j_in, "j_use": j_use, "j_stk": j_stk
        })

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

        if g["y_in"] > 0: item_queues["yangsangchu"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["y_in"], "vendor": "K01"})
        if g["b_in"] > 0: item_queues["yangbaechu"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["b_in"], "vendor": "K01"})
        if g["j_in"] > 0: item_queues["jeokchae"].append({"yymmdd": yymmdd, "label": f"{mmdd_str} 매입", "avail": g["j_in"], "vendor": "K01"})

        def trace_usage(item_key, item_prefix, use_amount):
            if use_amount <= 0: return "-", "-"
            rem = use_amount
            used_sources, used_codes = [], []

            for lot in item_queues[item_key]:
                if rem <= 0: break
                avail_val = lot.get("avail", 0.0)
                if avail_val <= 0: continue
                
                u = min(avail_val, rem)
                lot["avail"] = avail_val - u
                rem -= u

                vendor = lot.get("vendor", "K01")
                used_sources.append(f"{lot['label']}({round(u/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{lot['yymmdd']}-{vendor}")

            if rem > 0:
                used_sources.append(f"부족({round(rem/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{yymmdd}-K01[부족]")

            return " / ".join(used_sources), " / ".join(list(dict.fromkeys(used_codes)))

        y_src, y_code = trace_usage("yangsangchu", "X", g["y_use"])
        b_src, b_code = trace_usage("yangbaechu", "D", g["b_use"])
        j_src, j_code = trace_usage("jeokchae", "A", g["j_use"])

        note = "휴무" if (g["y_in"] == 0 and g["y_use"] == 0 and g["b_in"] == 0 and g["b_use"] == 0 and g["j_in"] == 0 and g["j_use"] == 0) else ""

        row_data = {
            "date": dt, "yymmdd": yymmdd,
            "yangsangchu_in": round(g["y_in"] / 1000.0, 1), 
            "yangsangchu_use": round(g["y_use"] / 1000.0, 1), 
            "yangsangchu_source": y_src, 
            "yangsangchu_code": y_code,
            "yangsangchu_stock": round(g["y_stk"] / 1000.0, 1),
            "yangbaechu_in": round(g["b_in"] / 1000.0, 1), 
            "yangbaechu_use": round(g["b_use"] / 1000.0, 1), 
            "yangbaechu_source": b_src, 
            "yangbaechu_code": b_code,
            "yangbaechu_stock": round(g["b_stk"] / 1000.0, 1),
            "jeokchae_in": round(g["j_in"] / 1000.0, 1), 
            "jeok
