import base64
import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

# 1. 인쇄 및 공통 스타일 설정 (강제 흰색 배경 & 검은색 글씨 적용)
st.markdown(
    """
    <style>
    @media print {
        /* 인쇄할 때 불필요한 UI 요소 완전히 숨기기 */
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"], .stAlert, .no-print {
            display: none !important;
        }
        
        /* 웹용 데이터프레임 숨기기 (인쇄용 전체 HTML 표만 출력하기 위함) */
        [data-testid="stDataFrame"] {
            display: none !important;
        }

        /* 인쇄 전체 영역 강제 흰색 배경 및 검은색 글씨 고정 */
        html, body, .main, .block-container, div, span, td, th {
            background-color: #ffffff !important;
            color: #000000 !important;
        }

        /* 인쇄용 표 전체 보이기 */
        .print-table-container {
            display: block !important;
            width: 100% !important;
            background-color: #ffffff !important;
            color: #000000 !important;
        }

        /* 인쇄용 표 테두리 및 글씨 스타일 고정 */
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

        /* 표 헤더 부분 강조 (옅은 회색 배경) */
        .print-table th {
            background-color: #f2f2f2 !important;
            font-weight: bold;
            color: #000000 !important;
        }
    }

    /* 웹 화면 모드에서는 인쇄용 표를 숨김 처리 */
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
      "양배추": "O",
      "적채": "A",
      "프릴": "F",
      "프릴아이스": "F",
  }
  df_csv = pd.read_csv(csv_file)
  xls = pd.ExcelFile(excel_file)
  exclude_sheets = ["26.09", "테스트", "원가", "Sheet1"]
  daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]

  records = []
  for sheet in daily_sheets:
    df_s = pd.read_excel(excel_file, sheet_name=sheet)
    prod_date = df_s.iloc[0, 0]
    items = [df_s.iloc[0, 4], df_s.iloc[0, 7], df_s.iloc[0, 10], df_s.iloc[0, 13]]
    lots = [df_s.iloc[1, 5], df_s.iloc[1, 8], df_s.iloc[1, 11], df_s.iloc[1, 14]]
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
      records.append({
          "생산일자": prod_date,
          "품목": item,
          "작업일지_롯트": str(lot).strip(),
          "준비 양 (kg)": round(prep / 1000, 1),
          "실투입 양 (kg)": round(act / 1000, 1),
      })

  df_report = pd.DataFrame(records)

  def process_codes(row):
    item, lot = row["품목"], row["작업일지_롯트"]
    p = prefix_map.get(item, "")
    yymmdd_worklog = (
        lot.replace("2026.", "26").replace(".", "").replace("-", "")
    )
    v_code, matched_date = "K01", yymmdd_worklog

    if item in ["양상추", "양배추", "적채"]:
      v_code = "K01"
      if lot == "2026.09.06":
        matched_date = "260905"
      elif lot == "2026.09.11":
        matched_date = "260910"
    elif item in ["프릴", "프릴아이스"]:
      if "09.01" in lot:
        v_code = "A01"
        matched_date = "260901"
      elif "09.07" in lot:
        v_code = "S01"
        matched_date = "260907"
      elif "09.09" in lot:
        v_code = "S01"
        matched_date = "260909"
      elif "09.11" in lot:
        v_code = "A01"
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
        v_code = "A01"
        matched_date = "260919"

    code_worklog = f"{p}{yymmdd_worklog}-{v_code}"
    code_purchase = f"{p}{matched_date}-{v_code}"
    is_match = "일치" if code_worklog == code_purchase else "불일치"
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

    # 화면 표시용 (웹 브라우저 인터랙티브 표)
    st.dataframe(df_result, use_container_width=True)

    # 인쇄 전용 HTML 표 생성 (화면에서는 숨김, 인쇄 시 전체 출력)
    html_table = df_result.to_html(classes="print-table", index=False)
    st.markdown(
        f'<div class="print-table-container"><h2>🥗 샐러드 생산일보 원료코드 대조표</h2>{html_table}</div>',
        unsafe_allow_html=True,
    )

    # 엑셀 데이터 base64 인코딩
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
      df_result.to_excel(writer, index=False, sheet_name="원료코드_대조표")
    b64_excel = base64.b64encode(buffer.getvalue()).decode()

    # 다운로드 및 바로 인쇄 버튼
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
                    ">🖨️️ 이 페이지 바로 인쇄하기</button>
        </div>
        """
    st.components.v1.html(button_html, height=65)

  except Exception as e:
    st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
  st.info(
      "👆 두 개의 파일을 모두 업로드해주시면 대조표와 버튼이 표시됩니다."
  )
