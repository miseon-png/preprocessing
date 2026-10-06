import base64
import io
import re
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
    # 품목별 코드 접두사
    prefix_map = {
        "양상추": "X",
        "양배추": "D",
        "적채": "A",
        "프릴": "F",
        "프릴아이스": "F",
    }

    # 거래처별 공급업체 코드
    vendor_code_map = {
        "승승장구": "S01",
        "애상스팜": "H02",
        "한스": "K01",
    }

    # CSV 매입 데이터 읽기 및 컬럼 정리
    df_csv = pd.read_csv(csv_file)
    df_csv.columns = [str(col).strip() for col in df_csv.columns]

    # 주요 컬럼 자동 감지
    date_col = next((c for c in df_csv.columns if "일자" in c or "날짜" in c), None)
    item_col = next((c for c in df_csv.columns if "품목" in c or "원재료" in c or "품명" in c), None)
    vendor_col = next((c for c in df_csv.columns if "거래처" in c or "공급" in c or "매입처" in c), None)

    # 날짜 표준화 함수 (YYMMDD)
    def normalize_date(val):
        if pd.isna(val):
            return ""
        digits = re.sub(r"\D", "", str(val))
        if len(digits) == 8:    # 20260909 -> 260909
            return digits[2:]
        elif len(digits) == 6:  # 260909
            return digits
        return digits

    if date_col:
        df_csv["_norm_date"] = df_csv[date_col].apply(normalize_date)

    xls = pd.ExcelFile(excel_file)
    exclude_sheets = ["26.09", "테스트", "원가", "Sheet1"]
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]

    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet)
        prod_date = df_s.iloc[0, 0]  # 생산일자

        # -------------------------------------------------------------
        # 품목별 좌표 매핑 (0-based index)
        # 2행(1): 품목명 / 3행(2): 롯트 / 4행(3): 준비양 / 6행(5): 실투입양
        # -------------------------------------------------------------
        items_config = [
            # 양상추: E2, F3, E4, E6
            (df_s.iloc[1, 4], df_s.iloc[2, 5], df_s.iloc[3, 4], df_s.iloc[5, 4]),
            # 양배추: H2, I3, H4, H6
            (df_s.iloc[1, 7], df_s.iloc[2, 8], df_s.iloc[3, 7], df_s.iloc[5, 7]),
            # 적채: K2, L3, K4, K6
            (df_s.iloc[1, 10], df_s.iloc[2, 11], df_s.iloc[3, 10], df_s.iloc[5, 10]),
            # 프릴: N2, O3, N4, N6
            (df_s.iloc[1, 13], df_s.iloc[2, 14], df_s.iloc[3, 13], df_s.iloc[5, 13]),
        ]

        for item, lot, prep, act in items_config:
            if pd.isna(item) or str(item).strip() == "":
                continue
            records.append({
                "생산일자": prod_date,
                "품목": str(item).strip(),
                "작업일지_롯트": str(lot).strip() if pd.notna(lot) else "",
                "준비 양 (kg)": round(prep / 1000, 1) if pd.notna(prep) and isinstance(prep, (int, float)) else 0.0,
                "실투입 양 (kg)": round(act / 1000, 1) if pd.notna(act) and isinstance(act, (int, float)) else 0.0,
            })

    df_report = pd.DataFrame(records)

    def process_codes(row):
        item = row["품목"]
        lot = row["작업일지_롯트"]

        # '프릴'은 매입 데이터의 '프릴아이스'로 동일 매칭
        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        # 작업일지 롯트 날짜 (YYMMDD)
        worklog_date = normalize_date(lot)

        # 작업일지 기준 기본 거래처 추정
        default_vcode = "K01"  # 기본 한스
        if "프릴" in item:
            if any(k in lot for k in ["09.01", "09.11", "09.19"]):
                default_vcode = "H02"  # 애상스팜
            elif any(k in lot for k in ["09.07", "09.09", "09.12", "09.16", "09.17"]):
                default_vcode = "S01"  # 승승장구

        code_worklog = f"{p}{worklog_date}-{default_vcode}" if worklog_date else "롯트미입력"

        # CSV 매입 데이터와 실제 대조
        code_purchase = "매입내역없음"

        if date_col and item_col and vendor_col and worklog_date:
            # 품목 및 롯트 날짜 기준 매입 내역 검색
            matched_rows = df_csv[
                (df_csv[item_col].astype(str).str.strip().isin([item, lookup_item])) &
                (df_csv["_norm_date"] == worklog_date)
            ]

            # 예외 롯트 날짜 보정 재검색 (9월 12일 롯트 -> 9월 9일 입고분)
            if matched_rows.empty and "0912" in worklog_date:
                matched_rows = df_csv[
                    (df_csv[item_col].astype(str).str.strip().isin([item, lookup_item])) &
                    (df_csv["_norm_date"] == "260909")
                ]

            # 매입 내역이 실제로 확인된 경우
            if not matched_rows.empty:
                vendor_name = str(matched_rows.iloc[0][vendor_col]).strip()
                matched_vcode = vendor_code_map.get(vendor_name, "K01")
                matched_date = matched_rows.iloc[0]["_norm_date"]
                code_purchase = f"{p}{matched_date}-{matched_vcode}"

        # 일치여부 최종 검증
        is_match = "일치" if code_worklog == code_purchase and code_worklog != "롯트미입력" else "불일치"

        return pd.Series([code_worklog, code_purchase, is_match])

    df_report[["작업일지 기준 코드", "매입자료 기준 코드", "일치여부"]] = (
        df_report.apply(process_codes, axis=1)
    )

    return df_report[[
        "생산일자",
        "품목",
        "작업일지 기준 코드",
        "매입자료 기준 코드",
        "일치여부",
        "준비 양 (kg)",
        "실투입 양 (kg)",
    ]]


if excel_file is not None and csv_file is not None:
    try:
        df_result = generate_salad_report(excel_file, csv_file)
        st.success("✅ 대조가 완료되었습니다!")

        st.dataframe(df_result, use_container_width=True)

        html_table = df_result.to_html(classes="print-table", index=False)
        st.markdown(
            f'<div class="print-table-container"><h2>🥗 샐러드 생산일보 원료코드 대조표</h2>{html_table}</div>',
            unsafe_allow_html=True,
        )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="원료코드_대조표")
        b64_excel = base64.b64encode(buffer.getvalue()).decode()

        button_html = f"""
            <div style="display: flex; gap: 16px; width: 100%; margin-top: 8px;">
                <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" 
                   download="원료코드_비교_정리표.xlsx" 
                   style="
                       flex: 1;
                       height: 48px;
                       background-color: #2E7D32;
                       color: white;
                       text-decoration: none;
                       display: flex;
                       align-items: center;
                       justify-content: center;
                       font-size: 16px;
                       font-weight: bold;
                       border-radius: 8px;
                       box-sizing: border-box;
                   ">📥 엑셀 파일 다운로드 (.xlsx)</a>
                <button onclick="window.parent.print()" 
                        style="
                            flex: 1;
                            height: 48px;
                            background-color: #2E7D32;
                            color: white;
                            border: none;
                            display: flex;
                            align-items: center;
                            justify-content: center;
                            font-size: 16px;
                            font-weight: bold;
                            border-radius: 8px;
                            cursor: pointer;
                            box-sizing: border-box;
                        ">🖨 이 페이지 바로 인쇄하기</button>
            </div>
            """
        st.components.v1.html(button_html, height=65)

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info(
        "👆 두 개의 파일을 모두 업로드해주시면 대조표와 버튼이 표시됩니다."
    )
