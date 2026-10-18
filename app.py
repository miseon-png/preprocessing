def parse_stock_sheet_fifo(xls):
    """
    재고 시트 파싱 개선:
    - 날짜 행의 '양상추/양배추/적채' 컬럼 숫자는 [투입 후 기말재고]이므로 입고 로트로 추가하지 않음.
    - 첫번째 행(전월이월 재고)과 '입고'가 포함된 행만 실제 신규 입고 로트로 큐(Queue)에 추가함.
    """
    item_lots = {item: [] for item in item_col_map}
    try:
        stock_sheet = find_stock_sheet_name(xls)
        df_stock = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

        # 헤더 아래 실제 데이터 시작 행 탐색
        for idx, row in df_stock.iterrows():
            a_val = (
                str(row.iloc[0]).strip()
                if len(row) > 0 and pd.notna(row.iloc[0])
                else ""
            )
            if not a_val or a_val.lower() == "nan" or "매입" in a_val:
                continue

            yy, yymmdd, mmdd = extract_yymmdd_and_mmdd(a_val)
            is_incoming = "입고" in a_val

            # 첫 번째 데이터 행(기초 이월 재고) 판단
            is_first_data_row = (idx == 1 or (idx > 0 and "2026-09" in a_val))

            for item, cols in item_col_map.items():
                s_col = cols["stock"]
                s_val = (
                    row.iloc[s_col]
                    if s_col < len(row) and pd.notna(row.iloc[s_col])
                    else None
                )

                try:
                    cleaned_val = re.sub(r"[^\d.]", "", str(s_val))
                    s_num = float(cleaned_val) if cleaned_val else 0.0
                except Exception:
                    s_num = 0.0

                # ★ 핵심 수정: '입고' 행이거나 '첫번째 기초재고 행'일 때만 신규 로트로 등록 (단위: g -> kg)
                if (is_incoming or is_first_data_row) and s_num > 0:
                    item_lots[item].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(s_num / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )

        return item_lots, stock_sheet
    except Exception as e:
        st.warning(f"재고 시트 파싱 중 알림: {e}")
        return item_lots, xls.sheet_names[0]
