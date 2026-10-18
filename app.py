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

    # --- FIFO 기반 투입 소진 출처 상세 추적 및 코드 조합 생성 ---
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

        # 원료코드 생성탭 전용 행 수집
        if row["yangsangchu_use"] > 0 or row["yangbaechu_use"] > 0 or row["jeokchae_use"] > 0:
            generated_code_rows.append(
                {
                    "생산일자": row["date"],
                    "양상추 투입(kg)": round(row["yangsangchu_use"] / 1000.0, 1),
                    "양상추 소진매입일": y_src,
                    "양상추 생성코드": y_code if row["yangsangchu_use"] > 0 else "-",
                    "양배추 투입(kg)": round(row["yangbaechu_use"] / 1000.0, 1),
                    "양배추 소진매입일": b_src,
                    "양배추 생성코드": b_code if row["yangbaechu_use"] > 0 else "-",
                    "적채 투입(kg)": round(row["jeokchae_use"] / 1000.0, 1),
                    "적채 소진매입일": j_src,
                    "적채 생성코드": j_code if row["jeokchae_use"] > 0 else "-",
                    "비고": row["note"],
                }
            )

    # --- openpyxl 엑셀 생성 ---
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
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = regular_font
            cell.border = thin_border
            if c_idx in [2, 3, 4, 6, 7, 8, 9, 11, 12, 13, 14, 16]:
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")

            if is_rest:
                cell.fill = rest_fill
            elif r_idx % 2 == 1:
                cell.fill = zebra_fill

        ws.row_dimensions[r_idx].height = 20

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 13
    ws.column_dimensions["E"].width = 28
    ws.column_dimensions["J"].width = 28
    ws.column_dimensions["O"].width = 28
    ws.column_dimensions["Q"].width = 16

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue(), data_rows, generated_code_rows, stock_sheet


def create_lots_from_standard_ledger(standard_rows):
    item_lots = {"양상추": [], "양배추": [], "적채": []}

    if not standard_rows:
        return item_lots

    first_row = standard_rows[0]
    first_yymmdd = first_row["yymmdd"]

    if first_row["yangsangchu_beg"] > 0:
        item_lots["양상추"].append({"yymmdd": first_yymmdd, "avail_g": first_row["yangsangchu_beg"], "vendor": "K1"})
    if first_row["yangbaechu_beg"] > 0:
        item_lots["양배추"].append({"yymmdd": first_yymmdd, "avail_g": first_row["yangbaechu_beg"], "vendor": "K1"})
    if first_row["jeokchae_beg"] > 0:
        item_lots["적채"].append({"yymmdd": first_yymmdd, "avail_g": first_row["jeokchae_beg"], "vendor": "K1"})

    for row in standard_rows:
        yymmdd = row["yymmdd"]
        if row["yangsangchu_in"] > 0:
            item_lots["양상추"].append({"yymmdd": yymmdd, "avail_g": row["yangsangchu_in"], "vendor": "K1"})
        if row["yangbaechu_in"] > 0:
            item_lots["양배추"].append({"yymmdd": yymmdd, "avail_g": row["yangbaechu_in"], "vendor": "K1"})
        if row["jeokchae_in"] > 0:
            item_lots["적채"].append({"yymmdd": yymmdd, "avail_g": row["jeokchae_in"], "vendor": "K1"})

    return item_lots


def parse_csv_purchase(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file)
        except Exception:
            return pd.DataFrame()

    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except Exception:
            continue

    return pd.DataFrame()


def get_latest_purchase_code(item_name, prod_date, df_csv):
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, prefix_map.get(str(item_name).strip(), "X"))
    _, p_yymmdd, p_mmdd = extract_yymmdd_and_mmdd(prod_date)
    v_code = "F01" if p_code == "F" else "K1"

    if isinstance(df_csv, pd.DataFrame) and not df_csv.empty:
        try:
            df_filtered = df_csv[
                df_csv.astype(str).apply(
                    lambda r: r.str.contains(lookup_item, regex=False).any(),
                    axis=1,
                )
            ]

            if not df_filtered.empty:
                candidate_rows = []
                for _, r in df_filtered.iterrows():
                    row_str = " ".join(r.dropna().astype(str).tolist())
                    _, r_yymmdd, r_mmdd = extract_yymmdd_and_mmdd(row_str)

                    if r_yymmdd <= p_yymmdd:
                        candidate_rows.append((r_yymmdd, r_mmdd, row_str))

                if candidate_rows:
                    candidate_rows.sort(key=lambda x: x[0])
                    matched_yymmdd = candidate_rows[-1][0]
                    row_str = candidate_rows[-1][2]

                    vendor_find = re.search(r"\b([A-Z]\d{2})\b", row_str)
                    if vendor_find:
                        v_code = vendor_find.group(1)

                    return f"{p_code}{matched_yymmdd}-{v_code}"
        except Exception:
            pass

    return f"{p_code}{p_yymmdd}-{v_code}"


def generate_salad_report(excel_file, csv_file, standard_rows, stock_sheet_name):
    xls = pd.ExcelFile(excel_file)

    master_item_lots = create_lots_from_standard_ledger(standard_rows)
    running_item_lots = copy.deepcopy(master_item_lots)

    df_csv = parse_csv_purchase(csv_file)

    exclude_sheets = ["테스트", "원가", stock_sheet_name, "Sheet1", "sheet1"]
    daily_sheets = [
        s
        for s in xls.sheet_names
        if s not in exclude_sheets and not s.startswith("재고")
    ]

    records = []
    for sheet in daily_sheets:
        try:
            df_s = pd.read_excel(excel_file, sheet_name=sheet)

            if df_s.shape[0] < 5 or df_s.shape[1] < 14:
                continue

            prod_date = df_s.iloc[0, 0]

            items = [
                df_s.iloc[0, 4],
                df_s.iloc[0, 7],
                df_s.iloc[0, 10],
                df_s.iloc[0, 13],
            ]
            lots = [
                df_s.iloc[1, 5],
                df_s.iloc[1, 8],
                df_s.iloc[1, 11],
                df_s.iloc[1, 14],
            ]
            prep_pck = [
                df_s.iloc[2, 4],
                df_s.iloc[2, 7],
                df_s.iloc[2, 10],
                df_s.iloc[2, 13],
            ]
            actual_input = [
                df_s.iloc[4, 4],
                df_s.iloc[4, 7],
                df_s.iloc[4, 10],
                df_s.iloc[4, 13],
            ]

            for item, lot, prep, act in zip(items, lots, prep_pck, actual_input):
                if (
                    pd.isna(item)
                    or str(item).strip() == ""
                    or str(item).strip().lower() == "nan"
                ):
                    continue

                try:
                    p_val = float(prep) if pd.notna(prep) else 0.0
                except (ValueError, TypeError):
                    p_val = 0.0

                try:
                    a_val = float(act) if pd.notna(act) else 0.0
                except (ValueError, TypeError):
                    a_val = 0.0

                lot_str = str(lot).strip() if pd.notna(lot) else ""
                has_lot = (
                    lot_str != ""
                    and lot_str.lower() != "nan"
                    and lot_str.lower() != "none"
                )

                _, prod_yymmdd, _ = extract_yymmdd_and_mmdd(prod_date)

                records.append(
                    {
                        "생산일자": prod_date,
                        "prod_yymmdd": prod_yymmdd,
                        "품목": str(item).strip(),
                        "작업일지_롯트": lot_str if has_lot else "없음",
                        "act_g": a_val,
                        "실투입 양 (kg)": round(a_val / 1000.0, 1),
                        "준비 양 (kg)": round(p_val / 1000.0, 1),
                        "has_lot": has_lot,
                    }
                )
        except Exception:
            continue

    if not records:
        st.error("생산일지 시트에서 처리 가능한 데이터를 찾지 못했습니다.")
        return pd.DataFrame()

    records.sort(key=lambda x: x["prod_yymmdd"])

    final_rows = []
    for rec in records:
        item = rec["품목"]
        lot = rec["작업일지_롯트"]
        has_lot = rec["has_lot"]
        prod_date = rec["생산일자"]
        prod_yymmdd = rec["prod_yymmdd"]
        act_g = rec["act_g"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        if has_lot:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = (
                f"{p}{yymmdd_worklog}-K1" if yymmdd_worklog else "없음"
            )
        else:
            code_worklog = "없음"

        code_subul = get_latest_purchase_code(item, prod_date, df_csv)

        if lookup_item not in master_item_lots:
            code_purchase = ""
        else:
            all_lots = running_item_lots.get(lookup_item, [])
            valid_lots = [l for l in all_lots if l["yymmdd"] <= prod_yymmdd]

            if not valid_lots or act_g <= 0:
                code_purchase = f"{p}{prod_yymmdd}-K1({round(act_g / 1000.0, 1)}kg)"
            else:
                remaining = act_g
                used_parts = []

                for lot_item in valid_lots:
                    if remaining <= 0:
                        break
                    if lot_item["avail_g"] <= 0:
                        continue

                    used = min(lot_item["avail_g"], remaining)
                    lot_item["avail_g"] -= used
                    remaining -= used

                    code_str = f"{p}{lot_item['yymmdd']}-{lot_item.get('vendor', 'K1')}"
                    used_parts.append(f"{code_str}({round(used / 1000.0, 1)}kg)")

                if remaining > 0:
                    latest_yymmdd = valid_lots[-1]["yymmdd"] if valid_lots else prod_yymmdd
                    v_code = valid_lots[-1].get("vendor", "K1") if valid_lots else "K1"
                    code_str = f"{p}{latest_yymmdd}-{v_code}"
                    used_parts.append(f"{code_str}({round(remaining / 1000.0, 1)}kg)[부족]")

                code_purchase = " / ".join(used_parts)

        if code_worklog != "없음" and code_purchase != "" and code_worklog in code_purchase:
            is_match = "일치"
        elif code_worklog != "없음" and code_purchase == "" and code_worklog == code_subul:
            is_match = "일치"
        elif code_worklog == "없음":
            is_match = "불일치(작업일지 미기재)"
        else:
            is_match = "불일치(혼합/추측)"

        rec["작업일지 기준 코드"] = code_worklog
        rec["수불부 기준 코드"] = code_subul
        rec["매입자료 기준 코드"] = code_purchase
        rec["일치여부"] = is_match
        final_rows.append(rec)

    df_report = pd.DataFrame(final_rows)

    return df_report[
        [
            "생산일자",
            "품목",
            "작업일지 기준 코드",
            "수불부 기준 코드",
            "매입자료 기준 코드",
            "일치여부",
            "실투입 양 (kg)",
            "준비 양 (kg)",
        ]
    ]


# ---------------------------------------------------------
# Streamlit 화면 구성 (3개 탭)
# ---------------------------------------------------------
tab1, tab2, tab3 = st.tabs(
    ["🔍 원료코드 대조표 (자동 연동)", "📊 표준 원재료 수불부 정리", "🏷️ 원료코드 생성 및 일자별 정리"]
)

if excel_file is not None:
    try:
        excel_bytes, standard_rows, code_summary_rows, stock_sheet_name = parse_and_convert_to_standard_ledger(excel_file)

        # --- Tab 1: 원료코드 대조표 ---
        with tab1:
            df_result = generate_salad_report(excel_file, csv_file, standard_rows, stock_sheet_name)
            if not df_result.empty:
                st.success("✅ 소진 추적이 연동된 표준 수불부 기반으로 원료코드 대조가 완료되었습니다!")

                st.dataframe(df_result, use_container_width=True)

                html_table = df_result.to_html(classes="print-table", index=False)
                st.markdown(
                    f'<div class="print-table-container"><h2>🥗 샐러드 생산일보 원료코드 대조표</h2>{html_table}</div>',
                    unsafe_allow_html=True,
                )

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    df_result.to_excel(writer, index=False, sheet_name="원료코드_대조표")

                st.download_button(
                    label="📥 원료코드 대조표 다운로드 (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="원료코드_비교_정리표.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption("💡 Tip: 인쇄를 원하시면 브라우저 인쇄 단축키(`Ctrl+P` 또는 `Cmd+P`)를 눌러주세요.")

        # --- Tab 2: 표준 원재료 수불부 ---
        with tab2:
            st.subheader("📋 표준 원재료 수불부 자동 생성 (투입 소진 매입일자 추적)")
            st.write(
                "기존 매입자료 시트를 **[일자 | 기초재고 | 입고량 | 투입량 | 소진 매입일자(상세) | 기말재고]** 구조의 깨끗한 표준 수불부 엑셀 문서로 정제해 드립니다."
            )

            df_preview = pd.DataFrame(standard_rows)
            if "yymmdd" in df_preview.columns:
                df_preview = df_preview.drop(columns=["yymmdd"])

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
                "yangsangchu_in": "양상추_입고량",
                "yangsangchu_use": "양상추_투입량",
                "yangsangchu_source": "양상추_소진 매입일자(상세)",
                "yangsangchu_stock": "양상추_기말재고",
                "yangbaechu_in": "양배추_입고량",
                "yangbaechu_use": "양배추_투입량",
                "yangbaechu_source": "양배추_소진 매입일자(상세)",
                "yangbaechu_stock": "양배추_기말재고",
                "jeokchae_in": "적채_입고량",
                "jeokchae_use": "적채_투입량",
                "jeokchae_source": "적채_소진 매입일자(상세)",
                "jeokchae_stock": "적채_기말재고",
                "note": "비고",
            }
            df_preview = df_preview.rename(columns=korean_column_map)

            st.write("▼ [입고량 ➔ 투입량 ➔ 소진 매입일자(상세) ➔ 기말재고] 수불부 미리보기")
            st.dataframe(df_preview, use_container_width=True)

            st.download_button(
                label="📥 정리된 표준 원재료 수불부 엑셀 다운로드 (.xlsx)",
                data=excel_bytes,
                file_name="표준_원재료_수불부(소진추적포함).xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # --- Tab 3: 신규 추가! 원료코드 생성 및 일자별 정리 ---
        with tab3:
            st.subheader("🏷️ 소진 매입재고 기반 일자별 원료코드 생성 결과")
            st.write(
                "**규칙**: 양상추(`X`), 양배추(`D`), 적채(`A`), 프릴(`F`) + `YYMMDD`(소진 매입일자) + `-K1` 거래처 코드 자동 부여"
            )

            df_code_summary = pd.DataFrame(code_summary_rows)
            if not df_code_summary.empty:
                st.dataframe(df_code_summary, use_container_width=True)

                code_buffer = io.BytesIO()
                with pd.ExcelWriter(code_buffer, engine="openpyxl") as writer:
                    df_code_summary.to_excel(writer, index=False, sheet_name="일자별_원료코드_생성")

                st.download_button(
                    label="📥 일자별 생성 원료코드 정리표 다운로드 (.xlsx)",
                    data=code_buffer.getvalue(),
                    file_name="일자별_생성_원료코드_목록.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            else:
                st.info("투입 수량 데이터가 존재하지 않습니다.")

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    with tab1:
        st.info("👆 상단에서 1. 샐러드 생산일지 엑셀 파일과 2. 원료 매입 수불부를 업로드해 주세요.")
    with tab2:
        st.info("👆 샐러드 생산일지 엑셀 파일을 업로드하시면 자동으로 정제된 수불부가 생성됩니다.")
    with tab3:
        st.info("👆 샐러드 생산일지 엑셀 파일을 업로드하시면 일자별로 자동 생성된 원료코드 표가 출력됩니다.")
