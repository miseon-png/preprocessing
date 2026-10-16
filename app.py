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

st.title("🥗 샐러드 생산일보 원료코드 대조 System")
excel_file = st.file_uploader(
    "샐러드 생산일지 (Excel) 업로드 (첫 시트: 당월 재고)", type=["xlsx", "xls"]
)

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}


def load_stock_data(xls):
    """첫 번째 시트에서 월간 재고 데이터를 읽어옵니다."""
    stock_sheet_name = xls.sheet_names[0]
    df_stock = pd.read_excel(xls, sheet_name=stock_sheet_name)

    # 컬럼명이 정형화되어 있지 않은 경우에 대비한 유연한 정제
    df_stock.dropna(how="all", inplace=True)
    return df_stock


def guess_purchase_code(item, prod_date, df_stock):
    """첫 번째 시트의 재고 내역을 참조하여 매입자료 기준 코드를 추측합니다."""
    if df_stock.empty:
        return "재고시트없음"

    lookup_item = "프릴아이스" if "프릴" in str(item) else str(item).strip()
    p_code = prefix_map.get(lookup_item, "X")

    # 재고 시트 내 검색 (품목 컬럼 대조)
    # 첫 시트에 품목, 입고일, 업체코드/원료코드 항목이 있을 경우 우선 탐색
    matched_rows = df_stock[
        df_stock.astype(str)
        .apply(lambda row: row.str.contains(lookup_item, regex=False))
        .any(axis=1)
    ]

    if not matched_rows.empty:
        # 최근 입고 내역 또는 첫 번째 매칭 재고의 코드 활용
        first_match = matched_rows.iloc[0]
        # 재고 시트에 추출할 수 있는 코드 형식이 있는 경우
        for val in first_match.values:
            val_str = str(val).strip()
            if re.match(r"^[A-Z]\d{6}-[A-Z0-9]+$", val_str):
                return val_str

    # 재고에서 개별 코드를 추출하지 못한 경우 생산일자 기반 추측 코드 생성 (예: X261017-K01)
    clean_date = re.sub(r"[^\d]", "", str(prod_date))
    yymmdd = clean_date[2:8] if len(clean_date) >= 8 else "261001"
    return f"{p_code}{yymmdd}-추측코드"


def generate_salad_report(excel_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 첫 번째 시트: 재고 데이터 로드
    df_stock = load_stock_data(xls)

    # 2. 일자별 생산 시트 순회 (첫 번째 시트 제외)
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [
        s
        for s in xls.sheet_names[1:]
        if s not in exclude_sheets and not s.startswith("재고")
    ]

    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet)

        # 생산일자 (A1)
        prod_date = df_s.iloc[0, 0]

        items = [
            df_s.iloc[0, 4],
            df_s.iloc[0, 7],
            df_s.iloc[0, 10],
            df_s.iloc[0, 13],
        ]  # E2, H2, K2, N2
        lots = [
            df_s.iloc[1, 5],
            df_s.iloc[1, 8],
            df_s.iloc[1, 11],
            df_s.iloc[1, 14],
        ]  # F3, I3, L3, O3
        prep_pck = [
            df_s.iloc[2, 4],
            df_s.iloc[2, 7],
            df_s.iloc[2, 10],
            df_s.iloc[2, 13],
        ]  # E4, H4, K4, N4
        actual_input = [
            df_s.iloc[4, 4],
            df_s.iloc[4, 7],
            df_s.iloc[4, 10],
            df_s.iloc[4, 13],
        ]  # E6, H6, K6, N6

        for item, lot, prep, act in zip(items, lots, prep_pck, actual_input):
            if (
                pd.isna(item)
                or str(item).strip() == ""
                or str(item).strip() == "nan"
            ):
                continue

            p_val = (
                float(prep)
                if pd.notna(prep) and isinstance(prep, (int, float))
                else 0.0
            )
            a_val = (
                float(act)
                if pd.notna(act) and isinstance(act, (int, float))
                else 0.0
            )

            # 롯트 번호 유효성 확인
            lot_str = str(lot).strip() if pd.notna(lot) else ""
            has_lot = (
                lot_str != ""
                and lot_str.lower() != "nan"
                and lot_str != "None"
            )

            records.append(
                {
                    "생산일자": prod_date,
                    "품목": str(item).strip(),
                    "작업일지_롯트": lot_str if has_lot else "없음",
                    "실투입 양 (kg)": round(a_val / 1000, 1),
                    "준비 양 (kg)": round(p_val / 1000, 1),
                    "has_lot": has_lot,
                }
            )

    df_report = pd.DataFrame(records)

    def process_codes(row):
        item = row["품목"]
        lot = row["작업일지_롯트"]
        has_lot = row["has_lot"]
        prod_date = row["생산일자"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        if has_lot:
            # 롯트번호가 있는 경우 연도/날짜 추출
            yymmdd_worklog = re.sub(r"[^\d]", "", lot)
            if len(yymmdd_worklog) >= 8:
                yymmdd_worklog = yymmdd_worklog[2:8]

            code_worklog = (
                f"{p}{yymmdd_worklog}-K01"
                if yymmdd_worklog
                else "코드변환실패"
            )
            # 재고 시트 참고하여 매입 코드 할당
            code_purchase = guess_purchase_code(item, prod_date, df_stock)
        else:
            # 롯트 번호가 없는 경우
            code_worklog = "없음"
            # 첫 번째 시트 재고를 바탕으로 추측 생성
            code_purchase = guess_purchase_code(item, prod_date, df_stock)

        is_match = (
            "일치"
            if code_worklog == code_purchase and code_worklog != "없음"
            else "불일치(추측)"
        )
        return pd.Series([code_worklog, code_purchase, is_match])

    df_report[["작업일지 기준 코드", "매입자료 기준 코드", "일치여부"]] = (
        df_report.apply(process_codes, axis=1)
    )

    return df_report[
        [
            "생산일자",
            "품목",
            "작업일지 기준 코드",
            "매입자료 기준 코드",
            "일치여부",
            "실투입 양 (kg)",
            "준비 양 (kg)",
        ]
    ]


if excel_file is not None:
    try:
        df_result = generate_salad_report(excel_file)
        st.success(
            "✅ 첫 시트(재고)와 일자별 시트 대조가 완료되었습니다!"
        )

        st.dataframe(df_result, use_container_width=True)

        html_table = df_result.to_html(classes="print-table", index=False)
        st.markdown(
            f'<div class="print-table-container"><h2>🥗 샐러드 생산일보 원료코드 대조표</h2>{html_table}</div>',
            unsafe_allow_html=True,
        )

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(
                writer, index=False, sheet_name="원료코드_대조표"
            )
        b64_excel = base64.b64encode(buffer.getvalue()).decode()

        button_html = (
            '<div style="display: flex; gap: 16px; width: 100%; margin-top: 8px;">'
            f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_excel}" '
            'download="원료코드_비교_정리표.xlsx" '
            'style="flex: 1; height: 48px; background-color: #2E7D32; color: white; text-decoration: none; '
            'display: flex; align-items: center; justify-content: center; font-size: 16px; font-weight: bold; '
            'border-radius: 8px; box-sizing: border-box;">📥 엑셀 파일 다운로드 (.xlsx)</a>'
            '<button onclick="window.parent.print()" '
            'style="flex: 1; height: 48px; background-color: #2E7D32; color: white; border: none; '
            'display: flex; align-items: center; justify-content: center; font-size: 16px; font-weight: bold; '
            'border-radius: 8px; cursor: pointer; box-sizing: border-box;">🖨 이 페이지 바로 인쇄하기</button>'
            "</div>"
        )
        st.components.v1.html(button_html, height=65)

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 샐러드 생산일지 엑셀 파일을 업로드해주세요.")
