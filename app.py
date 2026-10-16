import base64
import io
import re
import traceback
import pandas as pd
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

# 1. 인쇄 및 공통 스타일 설정
st.markdown(
    """
    <style>
    @media print {
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"], .stAlert, .no-print {
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
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table th, .print-table td {
            border: 1px solid #333333 !important;
            padding: 6px 4px;
            text-align: center;
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table th {
            background-color: #f2f2f2 !important;
            font-weight: bold;
            color: #000000 !important;
        }
    }
    .print-table-container {
        display: none;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("🥗 샐러드 생산일보 원료코드 대조 System")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "1. 샐러드 생산일지 (Excel) 업로드 (첫 시트: 당월 재고)",
        type=["xlsx", "xls"],
    )
with col2:
    csv_file = st.file_uploader(
        "2. 원료 매입 입고자료 (CSV/Excel) 업로드", type=["csv", "xlsx", "xls"]
    )

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}

# 품목별 첫 시트(재고 시트) 컬럼 인덱스 매핑 (A=0, B=1, D=3, F=5)
item_col_map = {
    "양상추": 1,  # B열
    "양배추": 3,  # D열
    "적채": 5,  # F열
}


def parse_stock_sheet(xls):
    """첫 시트(재고)에서 A열 입고일자와 B/D/F열 재고 정보 추출"""
    try:
        stock_sheet = xls.sheet_names[0]
        df_stock = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

        stock_records = []
        for idx, row in df_stock.iterrows():
            a_val = str(row.iloc[0]) if len(row) > 0 and pd.notna(row.iloc[0]) else ""

            # '10/05 입고' 또는 '10/05' 형태에서 MMDD 추출
            date_match = re.search(r"(\d{1,2})/(\d{1,2})", a_val)
            if date_match:
                mm, dd = date_match.group(1).zfill(2), date_match.group(2).zfill(2)
                mmdd = f"{mm}{dd}"

                row_info = {"mmdd": mmdd}
                for item_name, col_idx in item_col_map.items():
                    cell_val = row.iloc[col_idx] if col_idx < len(row) else None
                    has_stock = (
                        pd.notna(cell_val)
                        and str(cell_val).strip() not in ["", "0", "nan", "None"]
                    )
                    row_info[item_name] = has_stock

                stock_records.append(row_info)

        return stock_records
    except Exception as e:
        st.warning(f"재고 시트 해석 중 알림: {e}")
        return []


def parse_csv_purchase(uploaded_file):
    """CSV 또는 Excel 매입 입고자료 안전하게 읽기"""
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file)
        except Exception as e:
            st.warning(f"매입 엑셀 파일 읽기 실패: {e}")
            return pd.DataFrame()

    # CSV 인코딩 자동 우회 탐색
    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding=enc)
            return df
        except Exception:
            continue

    st.warning("매입 CSV 인코딩을 해석할 수 없습니다.")
    return pd.DataFrame()


def guess_purchase_code(item, prod_date, stock_records, df_csv):
    """
    생산일자 기준 매입자료 코드 추측
    """
    lookup_item = "프릴아이스" if "프릴" in str(item) else str(item).strip()
    p_code = prefix_map.get(lookup_item, "X")

    clean_pdate = re.sub(r"[^\d]", "", str(prod_date))
    year_prefix = clean_pdate[2:4] if len(clean_pdate) >= 8 else "26"
    p_mmdd = clean_pdate[4:8] if len(clean_pdate) >= 8 else ""

    v_code = "K01"

    # [케이스 1] 양상추, 양배추, 적채 -> 첫 시트 재고 참조
    if lookup_item in item_col_map:
        matched_mmdd = None
        if stock_records:
            valid_dates = [
                r["mmdd"]
                for r in stock_records
                if r.get(lookup_item) and (not p_mmdd or r["mmdd"] <= p_mmdd)
            ]
            if valid_dates:
                matched_mmdd = valid_dates[-1]

        if not matched_mmdd:
            matched_mmdd = p_mmdd if p_mmdd else "1001"

        return f"{p_code}{year_prefix}{matched_mmdd}-{v_code}"

    # [케이스 2] 프릴 / 프릴아이스 및 기타 -> CSV/Excel 매입 입고자료에서 최신 입고분 탐색
    else:
        if isinstance(df_csv, pd.DataFrame) and not df_csv.empty:
            try:
                # 품목명이 포함된 행 검색
                df_filtered = df_csv[
                    df_csv.astype(str)
                    .apply(
                        lambda r: r.str.contains(lookup_item, regex=False).any(),
                        axis=1,
                    )
                ]

                if not df_filtered.empty:
                    latest_row = df_filtered.iloc[-1]
                    row_str = " ".join(latest_row.astype(str).tolist())

                    date_find = re.search(r"20?(\d{2})[-./]?(\d{2})[-./]?(\d{2})", row_str)
                    if date_find:
                        matched_date = f"{date_find.group(1)}{date_find.group(2)}{date_find.group(3)}"
                    else:
                        matched_date = f"{year_prefix}{p_mmdd}" if p_mmdd else f"{year_prefix}1001"

                    vendor_find = re.search(r"\b([A-Z]\d{2})\b", row_str)
                    if vendor_find:
                        v_code = vendor_find.group(1)

                    return f"{p_code}{matched_date}-{v_code}"
            except Exception:
                pass

        return f"{p_code}{year_prefix}{p_mmdd if p_mmdd else '1001'}-{v_code}"


def generate_salad_report(excel_file, csv_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 첫 번째 시트 (재고 데이터) 파싱
    stock_records = parse_stock_sheet(xls)

    # 2. CSV 매입자료 파싱
    df_csv = parse_csv_purchase(csv_file)

    # 3. 일자별 생산 시트 순회 (첫 번째 시트 제외)
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [
        s
        for s in xls.sheet_names[1:]
        if s not in exclude_sheets and not s.startswith("재고")
    ]

    records = []
    for sheet in daily_sheets:
