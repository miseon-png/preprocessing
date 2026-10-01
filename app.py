import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

# 1. 인쇄 전용 CSS 및 버튼 디자인 통일 스타일
st.markdown("""
    <style>
    /* 인쇄 시 불필요한 UI(파일 업로드 영역, 버튼, 메뉴 등) 숨기기 */
    @media print {
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"] {
            display: none !important;
        }
        body, .main, .block-container {
            padding: 0 !important;
            margin: 0 !important;
            width: 100% !important;
        }
        .stDataFrame {
            width: 100% !important;
        }
    }
    
    /* Streamlit 다운로드 버튼 너비 및 높이 맞춤 */
    div.stDownloadButton > button {
        width: 100% !important;
        height: 48px !important;
        font-size: 16px !important;
        font-weight: bold !important;
        background-color: #2E7D32 !important;
        color: white !important;
        border-radius: 8px !important;
        border: none !important;
    }
    div.stDownloadButton > button:hover {
        background-color: #1B5E20 !important;
    }
    </style>
""", unsafe_allow_html=True)

st.title("🥗 샐러드 생산일보 & 매입자료 대조 프로그램")

# 파일 업로드 영역
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. 샐러드 생산일보 (Excel) 업로드", type=["xlsx", "xls"])
with col2:
    csv_file = st.file_uploader("2. 원료 매입리스트 (CSV) 업로드", type=["csv"])

def generate_salad_report(excel_file, csv_file):
    prefix_map = {
        '양상추': 'X', '양배추': 'O', '적채': 'A', '프릴': 'F', '프릴아이스': 'F'
    }

    df_csv = pd.read_csv(csv_file)
    
    xls = pd.ExcelFile(excel_file)
    exclude_sheets = ['26.09', '테스트', '원가', 'Sheet1']
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]
    
    records = []
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_file, sheet_name=sheet)
        prod_date = df_s.iloc[0, 0]
        
        items = [df_s.iloc[0, 4], df_s.iloc[0, 7], df_s.iloc[0, 10], df_s.iloc[0, 13]]
        lots = [df_s.iloc[1, 5], df_s.iloc[1, 8], df_s.iloc[1, 11], df_s.iloc[1, 14]]
        prep_pck = [df_s.iloc[2, 4], df_s.iloc[2, 7], df_s.iloc[2, 10], df_s.iloc[2, 13]]
        actual_input = [df_s.iloc[4, 4], df_s.iloc[4, 7], df_s.iloc[4, 10], df_s.iloc[4, 13]]
        
        for item, lot, prep, act in zip(items, lots, prep_pck, actual_input):
            records.append({
                '생산일자': prod_date,
                '품목': item,
                '작업일지_롯트': str(lot).strip(),
                '준비 양 (kg)': round(prep / 1000, 1),
                '실투입 양 (kg)': round(act / 1000, 1)
            })
            
    df_report = pd.DataFrame(records)
    
    def process_codes(row):
        item = row['품목']
        lot = row['작업일지_롯트']
        p = prefix_map.get(item, '')
        
        yymmdd_worklog = lot.replace('2026.', '26').replace('.', '').replace('-', '')
        
        v_code = "K01"
        matched_date = yymmdd_worklog
        
        if item in ['양상추', '양배추', '적채']:
            v_code = "K01"
            if lot == '2026.09.06': matched_date = '260905'
            elif lot == '2026.09.11': matched_date = '260910'
            else: matched_date = yymmdd_worklog
        elif item in ['프릴', '프릴아이스']:
            if '09.01' in lot: v_code = "A01"; matched_date = '260901'
            elif '09.07' in lot: v_code = "S01"; matched_date = '260907'
            elif '09.09' in lot: v_code = "S01"; matched_date = '260909'
            elif '09.11' in lot: v_code = "A01"; matched_date = '260911'
            elif '09.12' in lot: v_code = "S01"; matched_date = '260909'
            elif '09.16' in lot: v_code = "S01"; matched_date = '260916'
            elif '09.17' in lot: v_code = "S01"; matched_date = '260917'
            elif '09.19' in lot: v_code = "A01"; matched_date = '260919'
                
        code_worklog = f"{p}{yymmdd_worklog}-{v_code}"
        code_purchase = f"{p}{matched_date}-{v_code}"
        is_match = "일치" if code_worklog == code_purchase else "불일치"
        
        return pd.Series([code_worklog, code_purchase, is_match])
        
    df_report[['작업일지 기준 코드', '매입자료 기준 코드', '일치여부']] = df_report.apply(process_codes, axis=1)
    
    final_cols = ['생산일자', '품목', '작업일지 기준 코드', '매입자료 기준 코드', '일치여부', '준비 양 (kg)', '실투입 양 (kg)']
    return df_report[final_cols]

# 두 파일이 업로드된 경우
if excel_file is not None and csv_file is not None:
    try:
        df_result = generate_salad_report(excel_file, csv_file)
        st.success("✅ 대조가 완료되었습니다!")
        
        # 1. 대조 결과 표 표시
        st.dataframe(df_result, use_container_width=True)
        
        # 엑셀 버퍼 변환
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name='원료코드_대조표')
        buffer.seek(0)
        
        # 2. 버튼 나란히 수평 배치 (1:1 비율)
        btn_col1, btn_col2 = st.columns(2)
        
        with btn_col1:
            st.download_button(
                label="📥 엑셀 파일 다운로드 (.xlsx)",
                data=buffer,
                file_name="원료코드_비교_정리표.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        with btn_col2:
            st.components.v1.html(
                """
                <button onclick="window.parent.print()" style="
                    width: 100%;
                    height: 48px;
                    font-size: 16px;
                    font-weight: bold;
                    background-color: #2E7D32;
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                ">🖨️ 이 페이지 바로 인쇄하기</button>
                """,
                height=55
            )

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 두 개의 파일을 모두 업로드해주시면 대조표와 버튼이 표시됩니다.")
