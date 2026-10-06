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

    # 원형 코드 기준 exclude_sheets 원복
    xls = pd.ExcelFile(excel_file)
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]

    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet)

        # A1 생산일자
        prod_date = df_s.iloc[0, 0]

        # -----------------------------------------------------------------
        # 요청해주신 정확한 셀 위치 지정
        # 0-index 기준: E2=iloc[0,4], F3=iloc[1,5], E4=iloc[2,4], E6=iloc[4,4]
        # (pd.read_excel기본 읽기 시 헤더 1행 포함에 맞춘 iloc)
        # -----------------------------------------------------------------
        items = [df_s.iloc[0, 4], df_s.iloc[0, 7], df_s.iloc[0, 10], df_s.iloc[0, 13]]      # E2, H2, K2, N2 (품목명)
        lots = [df_s.iloc[1, 5], df_s.iloc[1, 8], df_s.iloc[1, 11], df_s.iloc[1, 14]]       # F3, I3, L3, O3 (롯트)
        prep_pck = [df_s.iloc[2, 4], df_s.iloc[2, 7], df_s.iloc[2, 10], df_s.iloc[2, 13]]   # E4, H4, K4, N4 (준비양(후))
        actual_input = [df_s.iloc[4, 4], df_s.iloc[4, 7], df_s.iloc[4, 10], df_s.iloc[4, 13]] # E6, H6, K6, N6 (실투입(전))

        for item, lot, prep, act in zip(items, lots, prep_pck, actual_input):
            if pd.isna(item) or str(item).strip() == "" or str(item).strip() == "nan":
                continue

            # 양 수치 변환
            p_val = float(prep) if pd.notna(prep) and isinstance(prep, (int, float)) else 0.0
            a_val = float(act) if pd.notna(act) and isinstance(act, (int, float)) else 0.0

            records.append({
                "생산일자": prod_date,
                "품목": str(item).strip(),
                "작업일지_롯트": str(lot).strip() if pd.notna(lot) else "",
                "준비 양 (kg)": round(p_val / 1000, 1),
                "실투입 양 (kg)": round(a_val / 1000, 1),
            })

    df_report = pd.DataFrame(records)

    def process_codes(row):
        item, lot = row["품목"], row["작업일지_롯트"]
        
        # 품목 접두사 (프릴 -> 프릴아이스 매칭)
        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        # YYMMDD 날짜 변환
        yymmdd_worklog = (
            lot.replace("2026.", "26").replace(".", "").replace("-", "").strip()
        )
        v_code, matched_date = "K01", yymmdd_worklog

        # 원형 코드 기준 예외 및 거래처 조건
        if item in ["양상추", "양배추", "적채"]:
            v_code = "K01"
            if lot == "2026.09.06":
                matched_date = "260905"
            elif lot == "2026.09.11":
                matched_date = "260910"
        elif item in ["프릴", "프릴아이스"]:
            if "09.01" in lot:
                v_code = "H02"  # 애상스팜
                matched_date = "260901"
            elif "09.07" in lot:
                v_code = "S01"  # 승승장구
                matched_date = "260907"
            elif "09.09" in lot:
                v_code = "S01"
                matched_date = "260909"
            elif "09.11" in lot:
                v_code = "H02"
                matched_date = "260911"
            elif "09.12" in lot:
                v_code = "S01"
                matched_date = "260909"
            elif "09.16" in lot:
                v_code = "S01"
                matched_date = "260916"
            elif "09.17" in lot:
                v_code = "S01"
                matched_date = "260917"
            elif "09.19" in lot:
                v_code = "H02"
                matched_date = "260919"

        code_worklog = f"{p}{yymmdd_worklog}-{v_code}" if yymmdd_worklog else "롯트미입력"
        code_purchase = f"{p}{matched_date}-{v_code}" if matched_date else "매입내역없음"
        
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

        # 화면 표시용
        st.dataframe(df_result, use_container_width=True)

        # 인쇄 전용 HTML 표
        html_table = df_result.to_html(classes="print-table", index=False)
        st.markdown(
            f'<div class="print-table-container"><h2>🥗 샐러드 생산일보 원료코드 대조표</h2>{html_table}</div>',
            unsafe_allow_html=True,
        )

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df_result.to_excel(writer, index=False, sheet_name="원료코드_대조표")
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
            '</div>'
        )
        st.components.v1.html(button_html, height=65)

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 두 개의 파일을 모두 업로드해주시면 대조표와 버튼이 표시됩니다.")
