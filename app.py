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


def extract_yymmdd_and
