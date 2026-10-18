import base64
import copy
import io
import re
import traceback
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 및 원료코드 System", layout="wide")

# 1. 인쇄 및 공통 스타일 설정
st.markdown(
    """
    <style>
    @media print {
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"], .stAlert, .no-print, [data-testid="stToolbar"] {
            display: none !important;
        }
        [data-testid="stDataFrame"] {
            display: none !important;
        }
        html, body, .main, .block-container, div, span, td, th {
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table-container {
            display: block !important;
            width: 100% !important;
        }
        .print-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 10px;
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table th, .print-table td {
            border: 1px solid #333333 !important;
            padding: 5px 4px;
            text-align: center;
            white-space: nowrap;
        }
        .print-table th {
            background-color: #f2f2f2 !important;
            font-weight: bold;
        }
    }
    .print-table-container {
        display: none;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🥗 샐러드 생산일보 원료코드 대조 및 수불부 연동 System")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "1. 샐러드 생산일지 (Excel) 업로드 (당월 재고/매입자료 포함)",
        type=["xlsx", "xls"],
    )
with col2:
    csv_file = st.file_uploader(
        "2. 원료 매입 입고자료 / 수불부 (CSV/Excel) 업로드",
        type=["csv", "xlsx", "xls"],
    )

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}


def extract_yymmdd_and_mmdd(val):
    if pd.isna(val):
        today = datetime.now()
        return today.strftime("%y"), today.strftime("%y%m%d"), today.strftime("%m%d")

    if isinstance(val, (pd.Timestamp, datetime)):
        yy = val.strftime("%y")
        yymmdd = val.strftime("%y%m%d")
        mmdd = val.strftime("%m%d")
        return yy, yymmdd, mmdd

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
            yy = dt.strftime("%y")
            yymmdd = dt.strftime("%y%m%d")
            mmdd = dt.strftime("%m%d")
            return yy, yymmdd, mmdd
    except Exception:
        pass

    current_yy = datetime.now().strftime("%y")

    m_short = re.search(r"(\d{1,2})[-./](\d{1,2})", val_str)
    if m_short:
        p1 = m_short.group(1).zfill(2)
        p2 = m_short.group(2).zfill(2)
        if int(p1) > 12:
            yy = p1
            mm = p2
            dd = "01"
        else:
            yy = current_yy
            mm = p1
            dd = p2
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
# 표준 수불부 정제 및 FIFO 투입 소진 추적 함수
# ---------------------------------------------------------
def parse_and_convert_to_standard_ledger(excel_file):
    xls = pd.ExcelFile(excel_file)
    stock_sheet = find_stock_sheet_name(xls)
    df_raw = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

    def parse_cell(val):
        if pd.isna(val):
            return 0.0
        try:
            c = re.sub(r"[^\d.]", "", str(val))
            return float(c) if c else 0.0
        except Exception:
            return 0.0

    pending_in = {"yangsangchu": 0.0, "yangbaechu": 0.0, "jeokchae": 0.0}
    raw_data_rows = []
    is_first_row = True

    for idx, row in df_raw.iterrows():
        if len(row) == 0 or pd.isna(row.iloc[0]):
            continue

        raw_date = row.iloc[0]
        a_val = str(raw_date).strip()
        if not a_val or a_val.lower() == "nan" or "날짜" in a_val or "매입" in a_val:
            continue

        if "입고" in a_val:
            pending_in["yangsangchu"] += parse_cell(row.iloc[1]) if len(row) > 1 else 0.0
            pending_in["yangbaechu"] += parse_cell(row.iloc[3]) if len(row) > 3 else 0.0
            pending_in["jeokchae"] += parse_cell(row.iloc[5]) if len(row) > 5 else 0.0
            continue

        _, yymmdd, _ = extract_yymmdd_and_mmdd(raw_date)
        formatted_date = f"20{yymmdd[:2]}-{yymmdd[2:4]}-{yymmdd[4:]}"

        y_stock = parse_cell(row.iloc[1]) if len(row) > 1 else 0.0
        y_use = parse_cell(row.iloc[2]) if len(row) > 2 else 0.0

        b_stock = parse_cell(row.iloc[3]) if len(row) > 3 else 0.0
        b_use = parse_cell(row.iloc[4]) if len(row) > 4 else 0.0

        j_stock = parse_cell(row.iloc[5]) if len(row) > 5 else 0.0
        j_use = parse_cell(row.iloc[6]) if len(row) > 6 else 0.0

        y_in = pending_in["yangsangchu"]
        b_in = pending_in["yangbaechu"]
        j_in = pending_in["jeokchae"]

        pending_in = {"yangsangchu": 0.0, "yangbaechu": 0.0, "jeokchae": 0.0}

        if is_first_row:
            y_beg = y_stock
            b_beg = b_stock
            j_beg = j_stock
            is_first_row = False
        else:
            y_beg = 0.0
            b_beg = 0.0
            j_beg = 0.0

        note = "휴무" if (y_in == 0 and y_use == 0 and b_in == 0 and b_use == 0 and j_in == 0 and j_use == 0 and y_stock == 0 and b_stock == 0 and j_stock == 0) else ""

        raw_data_rows.append(
            {
                "date": formatted_date,
                "yymmdd": yymmdd,
                "yangsangchu_beg": y_beg,
                "yangsangchu_in": y_in,
                "yangsangchu_use": y_use,
                "yangsangchu_stock": y_stock,
                "yangbaechu_beg": b_beg,
                "yangbaechu_in": b_in,
                "yangbaechu_use": b_use,
                "yangbaechu_stock": b_stock,
                "jeokchae_beg": j_beg,
                "jeokchae_in": j_in,
                "jeokchae_use": j_use,
                "jeokchae_stock": j_stock,
                "note": note,
            }
        )

    # --- FIFO 기반 투입 소진 출처 상세 추적 및 원료코드 생성 ---
    item_queues = {"yangsangchu": [], "yangbaechu": [], "jeokchae": []}
    if raw_data_rows:
        f_row = raw_data_rows[0]
        if f_row["yangsangchu_beg"] > 0:
            item_queues["yangsangchu"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["yangsangchu_beg"]})
        if f_row["yangbaechu_beg"] > 0:
            item_queues["yangbaechu"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["yangbaechu_beg"]})
        if f_row["jeokchae_beg"] > 0:
            item_queues["jeokchae"].append({"yymmdd": f_row["yymmdd"], "label": f"{f_row['yymmdd'][2:4]}/{f_row['yymmdd'][4:]} 기초", "avail": f_row["jeokchae_beg"]})

    data_rows = []
    generated_code_rows = []

    for row in raw_data_rows:
        row_copy = copy.deepcopy(row)
        mmdd_str = f"{row['yymmdd'][2:4]}/{row['yymmdd'][4:]}"

        if row["yangsangchu_in"] > 0:
            item_queues["yangsangchu"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["yangsangchu_in"]})
        if row["yangbaechu_in"] > 0:
            item_queues["yangbaechu"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["yangbaechu_in"]})
        if row["jeokchae_in"] > 0:
            item_queues["jeokchae"].append({"yymmdd": row["yymmdd"], "label": f"{mmdd_str} 매입", "avail": row["jeokchae_in"]})

        def trace_usage(item_key, item_prefix, use_amount):
            if use_amount <= 0:
                return "-", "-"
            rem = use_amount
            used_sources = []
            used_codes = []

            for lot in item_queues[item_key]:
                if rem <= 0:
                    break
                if lot["avail"] <= 0:
                    continue
                u = min(lot["avail"], rem)
                lot["avail"] -= u
                rem -= u

                used_sources.append(f"{lot['label']}({round(u/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{lot['yymmdd']}-K1")

            if rem > 0:
                used_sources.append(f"부족({round(rem/1000.0, 1)}kg)")
                used_codes.append(f"{item_prefix}{row['yymmdd']}-K1[부족]")

            return " / ".join(used_sources), " / ".join(list(dict.fromkeys(used_codes)))

        y_src, y_code = trace_usage("yangsangchu", "X", row["yangsangchu_use"])
        b_src, b_code = trace_usage("yangbaechu", "D", row["yangbaechu_use"])
        j_src, j_code = trace_usage("jeokchae", "A", row["jeokchae_use"])

        row_copy["yangsangchu_source"] = y_src
        row_copy["yangbaechu_source"] = b_src
        row_copy["jeokchae_source"] = j_src

        data_rows.append(row_copy)

        # 원료코드 생성 탭 전용 행 수집 (소진매입일 컬럼 완전히 제외)
        if row["yangsangchu_use"] > 0 or row["yangbaechu_use"] > 0 or row["jeokchae_use"] > 0:
            generated_code_rows.append(
                {
                    "생산일자": row["date"],
                    "양상추 투입(kg)": round(row["yangsangchu_use"] / 1000.0, 1),
                    "양상추 생성코드": y_code if row["yangsangchu_use"] > 0 else "-",
                    "양배추 투입(kg)": round(row["yangbaechu_use"] / 1000.0, 1),
                    "양배추 생성코드": b_code if row["yangbaechu_use"] > 0 else "-",
                    "적채 투입(kg)": round(row["jeokchae_use"] / 1000.0, 1),
                    "적채 생성코드": j_code if row["jeokchae_use"] > 0 else "-",
                    "비고": row["note"],
                }
            )

    # --- openpyxl 엑셀 파일 생성 ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "원재료 수불부"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:Q1")
    title_cell = ws["A1"]
    title_cell.value = "원재료 수불부 (Raw Material Inventory Ledger - 소진 추적 포함)"
    title_cell.font = Font(name="맑은 고딕", size=16, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    header_fill_main = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_fill_sub = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    zebra_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    rest_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    header_font_main = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")
    header_font_sub = Font(name="맑은 고딕", size=10, bold=True, color="1F4E79")
    regular_font = Font(name="맑은 고딕", size=10)

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    headers_main = [
        ("A2:A3", "일자"),
        ("B2:F2", "양상추"),
        ("G2:K2", "양배추"),
        ("L2:P2", "적채"),
        ("Q2:Q3", "비고"),
    ]
    for cell_range, title in headers_main:
        ws.merge_cells(cell_range)
        top_left = cell_range.split(":")[0]
        ws[top_left] = title

    for col_idx in range(1, 18):
        c = ws.cell(row=2, column=col_idx)
        c.fill = header_fill_main
        c.font = header_font_main
        c.alignment = Alignment(horizontal="center", vertical="center")

    sub_headers = ["기초재고", "입고량", "투입량", "소진 매입일자(상세)", "기말재고"]
    for start_col in [2, 7, 12]:
        for idx, sub in enumerate(sub_headers):
            c = ws.cell(row=3, column=start_col + idx, value=sub)
            c.fill = header_fill_sub
            c.font = header_font_sub
            c.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 22

    start_row = 4
    for r_idx, item in enumerate(data_rows, start=start_row):
        ws.cell(row=r_idx, column=1, value=item["date"]).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(row=r_idx, column=17, value=item.get("note", "")).alignment = Alignment(horizontal="left", vertical="center")

        items_map = [
            (2, item.get("yangsangchu_beg", 0), item.get("yangsangchu_in", 0), item.get("yangsangchu_use", 0), item.get("yangsangchu_source", "-")),
            (7, item.get("yangbaechu_beg", 0), item.get("yangbaechu_in", 0), item.get("yangbaechu_use", 0), item.get("yangbaechu_source", "-")),
            (12, item.get("jeokchae_beg", 0), item.get("jeokchae_in", 0), item.get("jeokchae_use", 0), item.get("jeokchae_source", "-")),
        ]

        for start_col, beg_val, in_val, use_val, source_str in items_map:
            end_col_let = get_column_letter(start_col + 4)

            if r_idx == start_row:
                ws.cell(row=r_idx, column=start_col, value=beg_val)
            else:
                ws.cell(row=r_idx, column=start_col, value=f"={end_col_let}{r_idx-1}")

            ws.cell(row=r_idx, column=start_col + 1, value=in_val)
            ws.cell(row=r_idx, column=start_col + 2, value=use_val)
            ws.cell(row=r_idx, column=start_col + 3, value=source_str).alignment = Alignment(horizontal="center", vertical="center")

            c_beg = get_column_letter(start_col)
            c_in = get_column_letter(start_col + 1)
            c_use = get_column_letter(start_col + 2)
            ws.cell(row=r_idx, column=start_col + 4, value=f"={c_beg}{r_idx}+{c_in}{r_idx}-{c_use}{r_idx}")

        is_rest = "휴무" in item.get("note", "")
        for c_idx in range(1, 18):
            cell = ws.cell(row=r_idx, column=c_
