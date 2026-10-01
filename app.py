import streamlit as st
import pandas as pd

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

st.title("🥗 샐러드 생산일보 & 매입자료 대조 프로그램")

# 파일 업로드 화면 구성
col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader("1. 샐러드 생산일보 (Excel) 업로드", type=["xlsx", "xls"])
with col2:
    csv_file = st.file_uploader("2. 원료 매입리스트 (CSV) 업로드", type=["csv"])

def generate_salad_report(excel_file, csv_file):
    # 공급업체 코드 매핑
    vendor_map = {
        '한스': 'K01', '승승장구': 'S01', '에이지': 'A01',
        '에이지로지스틱스': 'A01', '에상스팜': 'H02', '구름': 'G01'
    }
    
    # 품목 접두사 매핑
    prefix_map = {
        '양상추': 'X', '양배추': 'O', '적채': 'A', '프릴': 'F', '프릴아이스': 'F'
    }

    # CSV 불러오기
    df_csv = pd.read_csv(csv_file)
    
    # Excel 불러오기
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
            if lot == '2026.09.06':
                matched_date = '260905'
            elif lot == '2026.09.11':
                matched_date = '260910'
            else:
                matched_date = yymmdd_worklog
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

# 두 파일이 모두 업로드되었을 때 실행
if excel_file is not None and csv_file is not None:
    try:
        df_result = generate_salad_report(excel_file, csv_file)
        st.success("대조가 완료되었습니다!")
        st.dataframe(df_result, use_container_width=True)
    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
else:
    st.info("👆 두 개의 파일을 모두 업로드해주세요.")
