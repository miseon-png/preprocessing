import base64
import io
import re
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

    xls = pd.ExcelFile(excel_file)
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]

    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet)

        # A1 생산일자
        prod_date = df_s.iloc[0, 0]

        items = [df_s.iloc[0, 4], df_s.iloc[0, 7], df_s.iloc[0, 10], df_s.iloc[0, 13]]      # E2, H2, K2, N2
        lots = [df_s.iloc[1, 5], df_s.iloc[1, 8], df_s.iloc[1, 11], df_s.iloc[1, 14]]       # F3, I3, L3, O3
        prep_pck = [df_s.iloc[2, 4], df_s.iloc[2, 7], df_s.iloc[2, 10], df_s.iloc[2, 13]]   # E4, H4, K4, N4
        actual_input = [df_s.iloc[4, 4], df_s.iloc[4, 7], df_s.iloc[4, 10], df_s.iloc[4, 13]] # E6, H6, K6, N6

        for item, lot, prep, act in zip(items, lots, prep_pck, actual_input):
            if pd.isna(item) or str(item).strip() == "" or str(item).strip() == "nan":
                continue

            p_val = float(prep) if pd.notna(prep) and isinstance(prep, (int, float)) else 0.0
            a_val = float(act) if pd.notna(act) and isinstance(act, (int, float)) else 0.0

            records.append({
                "생산일자": prod_date,
                "품목": str(item).strip(),
                "작업일지_롯트": str(lot).strip() if pd.notna(lot) else "",
                "실투입 양 (kg)": round(a_val / 1000, 1),
                "준비 양 (kg)
