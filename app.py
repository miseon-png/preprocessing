import base64
import io
import re
import traceback
import pandas as pd
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

# 1. 인쇄 및 공통 스타일 설정 (강제 흰색 배경 & 검은색 글씨 적용)
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

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "1. 샐러드 생산일보 (Excel) 업로드", type=["xlsx", "xls"]
    )
with col2:
    csv_file = st.file_uploader("2. 원료 매입리스트 (CSV) 업로드", type=["csv"])


def generate_salad_report(excel_file, csv_file):
    prefix_map = {
        "양상추": "X",
        "양배추": "D",
        "적채": "A",
        "프릴": "F",
        "프릴아이스": "F",
    }

    vendor_code_map = {
        "승승장구": "S01",
        "애상스팜": "H02",
        "한스": "K01",
    }

    # CSV 데이터 읽기 (인코딩 대응)
    try:
        df_csv = pd.read_csv(csv_file, encoding="utf-8")
    except UnicodeDecodeError:
        csv_file.seek(0)
        df_csv = pd.read_csv(csv_file, encoding="cp949")

    df_csv.columns = [str(col).strip() for col in df_csv.columns]

    date_col = next((c for c in df_csv.columns if "일자" in c or "날짜" in c), None)
    item_col = next((c for c in df_csv.columns if "품목" in c or "원재료" in c or "품명" in c), None)
    vendor_col = next((c for c in df_csv.columns if "거래처" in c or "공급" in c or "매입처" in c), None)

    def normalize_date(val):
        if pd.isna(val):
            return ""
        digits = re.sub(r"\D", "", str(val))
        if len(digits) >= 8:
            return digits[2:8]
        elif len(digits) == 6:
            return digits
        return digits

    if date_col:
        df_csv["_norm_date"] = df_csv[date_col].apply(normalize_date)

    xls = pd.ExcelFile(excel_file)
    exclude_sheets = ["26.09", "테스트", "원가", "Sheet1"]
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]

    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet, header=None)

        if df_s.shape[0] < 6 or df_s.shape[1] < 15:
            continue

        prod_date = df_s.iloc[0, 0]

        def safe_get(r, c):
            val = df_s.iloc[r, c]
            return val if pd.notna(val) else ""
