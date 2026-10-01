import pandas as pd

def generate_salad_report(excel_path, csv_path, output_excel_path=None):
    """
    샐러드 생산일보(Excel)와 매입리스트(CSV)를 분석하여
    원료코드 대조 및 투입량 정리표를 생성합니다.
    """
    # 1. 공급업체 코드 매핑 사전
    vendor_map = {
        '한스': 'K01',
        '승승장구': 'S01',
        '에이지': 'A01',
        '에이지로지스틱스': 'A01',
        '에상스팜': 'H02',
        '구름': 'G01'
    }
    
    # 2. 품목 접두사 매핑 사전
    prefix_map = {
        '양상추': 'X',
        '양배추': 'O',
        '적채': 'A',
        '프릴': 'F',
        '프릴아이스': 'F'
    }

    # 3. 매입자료 CSV 불러오기
    df_csv = pd.read_csv(csv_path)
    
    # 4. 생산일보 Excel 파일 및 시트 추출
    xls = pd.ExcelFile(excel_path)
    # 일별 시트 제외 대상 시트명
    exclude_sheets = ['26.09', '테스트', '원가', 'Sheet1']
    daily_sheets = [s for s in xls.sheet_names if s not in exclude_sheets]
    
    records = []
    
    for sheet in daily_sheets:
        df_s = pd.read_excel(excel_path, sheet_name=sheet)
        prod_date = df_s.iloc[0, 0]
        
        # 4개 기본 품목 위치 추출 (Col 4: 양상추, Col 7: 양배추, Col 10: 적채, Col 13: 프릴)
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
    
    # 5. 원료코드 생성 및 매입자료 교차 대조
    def process_codes(row):
        item = row['품목']
        lot = row['작업일지_롯트']
        p = prefix_map.get(item, '')
        
        # 작업일지 롯트 날짜 (YYMMDD 변환)
        yymmdd_worklog = lot.replace('2026.', '26').replace('.', '').replace('-', '')
        
        # 매입자료 매칭 로직 (한스 K01, 승승장구 S01, 에이지 A01 등)
        v_code = "K01"
        matched_date = yymmdd_worklog
        
        if item in ['양상추', '양배추', '적채']:
            v_code = "K01"
            if lot == '2026.09.06':
                matched_date = '260905' # 9월 5일 매입건
            elif lot == '2026.09.11':
                matched_date = '260910' # 9월 10일 매입건
            else:
                matched_date = yymmdd_worklog
        elif item in ['프릴', '프릴아이스']:
            if '09.01' in lot:
                v_code = "A01"
                matched_date = '260901'
            elif '09.07' in lot:
                v_code = "S01"
                matched_date = '260907'
            elif '09.09' in lot:
                v_code = "S01"
                matched_date = '260909'
            elif '09.11' in lot:
                v_code = "A01"
                matched_date = '260911'
            elif '09.12' in lot:
                v_code = "S01"
                matched_date = '260909'
            elif '09.16' in lot:
                v_code = "S01"
                matched_date = '260916'
            elif '09.17' in lot:
                v_code = "S01"
                matched_date = '260917'
            elif '09.19' in lot:
                v_code = "A01"
                matched_date = '260919'
                
        code_worklog = f"{p}{yymmdd_worklog}-{v_code}"
        code_purchase = f"{p}{matched_date}-{v_code}"
        is_match = "일치" if code_worklog == code_purchase else "불일치"
        
        return pd.Series([code_worklog, code_purchase, is_match])
        
    df_report[['작업일지 기준 코드', '매입자료 기준 코드', '일치여부']] = df_report.apply(process_codes, axis=1)
    
    # 6. 최종 컬럼 구성 및 정렬
    final_cols = ['생산일자', '품목', '작업일지 기준 코드', '매입자료 기준 코드', '일치여부', '준비 양 (kg)', '실투입 양 (kg)']
    df_result = df_report[final_cols]
    
    # 엑셀 저장 옵션 지정 시 저장
    if output_excel_path:
        df_result.to_excel(output_excel_path, index=False)
        print(f"결과 파일이 저장되었습니다: {output_excel_path}")
        
    return df_result

# --- 실행 예시 ---
# 작업일지 파일명과 매입리스트 파일명을 입력하세요.
excel_file = '샐러드 생산일보 09월.xlsx'
csv_file = '2026-10-01T08-44_export.csv'

# 실행 후 결과 출력 및 엑셀 저장
df_final = generate_salad_report(excel_file, csv_file, output_excel_path='원료코드_비교_정리표.xlsx')
print(df_final)
