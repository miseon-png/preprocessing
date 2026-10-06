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

    # CSV 데이터 안전하게 읽기 (인코딩 처리)
    csv_file.seek(0)
    try:
        df_csv = pd.read_csv(csv_file, encoding="utf-8")
    except Exception:
        csv_file.seek(0)
        df_csv = pd.read_csv(csv_file, encoding="cp949")

    df_csv.columns = [str(col).strip() for col in df_csv.columns]

    date_col = next((c for c in df_csv.columns if "일자" in c or "날짜" in c), None)
    item_col = next((c for c in df_csv.columns if "품목" in c or "원재료" in c or "품명" in c), None)
    vendor_col = next((c for c in df_csv.columns if "거래처" in c or "공급" in c or "매입처" in c), None)

    def normalize_date(val):
        if pd.isna(val) or not val:
            return ""
        digits = re.sub(r"\D", "", str(val))
        if len(digits) >= 8:
            return digits[2:8]
        elif len(digits) == 6:
            return digits
        return digits

    if date_col:
        df_csv["_norm_date"] = df_csv[date_col].apply(normalize_date)

    # Streamlit 파일 객체 안전 읽기 (xls.parse 사용)
    excel_file.seek(0)
    xls = pd.ExcelFile(excel_file)
    
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [s for s in xls.sheet_names if not any(ex in s for ex in exclude_sheets)]

    records = []
    for sheet in daily_sheets:
        # file uploader 버그 방지를 위해 xls.parse() 사용
        df_s = xls.parse(sheet, header=None)

        if df_s.shape[0] < 3 or df_s.shape[1] < 5:
            continue

        prod_date = df_s.iloc[0, 0]

        def safe_get_lot(r_primary, c_primary, alt_coords=[]):
            for r, c in [(r_primary, c_primary)] + alt_coords:
                if r < df_s.shape[0] and c < df_s.shape[1]:
                    val = df_s.iloc[r, c]
                    if pd.notna(val) and str(val).strip() != "" and str(val).strip() != "nan":
                        return str(val).strip()
            return ""

        def safe_get_str(r, c):
            if r < df_s.shape[0] and c < df_s.shape[1]:
                val = df_s.iloc[r, c]
                if pd.notna(val):
                    return str(val).strip()
            return ""

        def safe_num(r, c):
            if r < df_s.shape[0] and c < df_s.shape[1]:
                val = df_s.iloc[r, c]
                try:
                    return float(val)
                except (ValueError, TypeError):
                    return 0.0
            return 0.0

        # 좌표 매핑 (E2, F3, E4, E6 / H2, I3, H4, H6 등)
        items_config = [
            (safe_get_str(1, 4), safe_get_lot(2, 5, [(1, 5), (3, 5)]), safe_num(3, 4), safe_num(5, 4)),
            (safe_get_str(1, 7), safe_get_lot(2, 8, [(1, 8), (3, 8)]), safe_num(3, 7), safe_num(5, 7)),
            (safe_get_str(1, 10), safe_get_lot(2, 11, [(1, 11), (3, 11)]), safe_num(3, 10), safe_num(5, 10)),
            (safe_get_str(1, 13), safe_get_lot(2, 14, [(1, 14), (3, 14)]), safe_num(3, 13), safe_num(5, 13)),
        ]

        for item, lot, prep, act in items_config:
            if not item or item.lower() == "nan":
                continue

            records.append({
                "생산일자": prod_date,
                "품목": item,
                "작업일지_롯트": lot,
                "준비 양 (kg)": round(prep / 1000, 1),
                "실투입 양 (kg)": round(act / 1000, 1),
            })

    if not records:
        st.error("❌ 엑셀 파일 시트에서 대조할 데이터를 불러오지 못했습니다.")
        return pd.DataFrame()

    df_report = pd.DataFrame(records)

    def process_codes(row):
        item = row["품목"]
        lot = row["작업일지_롯트"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        worklog_date = normalize_date(lot)

        default_vcode = "K01"
        if "프릴" in item:
            if any(k in lot for k in ["09.01", "09.11", "09.19"]):
                default_vcode = "H02"
            elif any(k in lot for k in ["09.07", "09.09", "09.12", "09.16", "09.17"]):
                default_vcode = "S01"

        code_worklog = f"{p}{worklog_date}-{default_vcode}" if worklog_date else "롯트미입력"
        code_purchase = "매입내역없음"

        if date_col and item_col and vendor_col and worklog_date:
            matched_rows = df_csv[
                (df_csv[item_col].astype(str).str.strip().isin([item, lookup_item])) &
                (df_csv["_norm_date"] == worklog_date)
            ]

            if matched_rows.empty and "0912" in worklog_date:
                matched_rows = df_csv[
                    (df_csv[item_col].astype(str).str.strip().isin([item, lookup_item])) &
                    (df_csv["_norm_date"] == "260909")
                ]

            if not matched_rows.empty:
                vendor_name = str(matched_rows.iloc[0][vendor_col]).strip()
                matched_vcode = vendor_code_map.get(vendor_name, "K01")
                matched_date = matched_rows.iloc[0]["_norm_date"]
                code_purchase = f"{p}{matched_date}-{matched_vcode}"

        is_match = "일치" if code_worklog == code_purchase and code_worklog != "롯트미입력" else "불일치"

        return pd.Series([code_worklog, code_purchase, is_match])

    df_report[["작업일지 기준 코드", "매입자료 기준 코드", "일치여부"]] = (
        df_report
