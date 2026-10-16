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

st.title("🥗 샐러드 생산일보 원료코드 대조 System")

col1, col2 = st.columns(2)
with col1:
    excel_file = st.file_uploader(
        "1. 샐러드 생산일지 (Excel) 업로드 (첫 시트: 당월 재고)",
        type=["xlsx", "xls"],
    )
with col2:
    csv_file = st.file_uploader(
        "2. 원료 매입 입고자료 (CSV/Excel) 업로드", type=["csv", "xlsx", "xls"]
    )

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}

# 품목별 첫 시트(재고 시트) 컬럼 인덱스 매핑 (A=0, B=1, D=3, F=5)
item_col_map = {
    "양상추": 1,  # B열
    "양배추": 3,  # D열
    "적채": 5,  # F열
}


def extract_yymmdd_and_mmdd(date_val):
    """다양한 형태의 날짜(26.10.05, 2026-10-05, 10/05 등)에서 YYMMDD 및 MMDD 추출"""
    date_str = str(date_val).strip()
    digits = re.sub(r"[^\d]", "", date_str)

    # 1) YYYYMMDD 형태 (8자리) -> ex: 20261005
    if len(digits) == 8:
        return digits[2:4], digits[2:8], digits[4:8]  # YY, YYMMDD, MMDD

    # 2) YYMMDD 형태 (6자리) -> ex: 261005
    elif len(digits) == 6:
        return digits[0:2], digits, digits[2:6]  # YY, YYMMDD, MMDD

    # 3) MMDD 형태 (4자리) -> ex: 1005
    elif len(digits) == 4:
        return "26", f"26{digits}", digits

    # 4) 정규표현식으로 월/일 추출 시도 (ex: 10/05, 10.5 등)
    match = re.search(r"(\d{1,2})[-./](\d{1,2})", date_str)
    if match:
        mm = match.group(1).zfill(2)
        dd = match.group(2).zfill(2)
        return "26", f"26{mm}{dd}", f"{mm}{dd}"

    return "26", "261001", "1001"


def parse_stock_sheet(xls):
    """첫 시트(재고)에서 A열 입고일자와 B/D/F열 재고 정보 추출"""
    try:
        stock_sheet = xls.sheet_names[0]
        df_stock = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

        stock_records = []
        for idx, row in df_stock.iterrows():
            a_val = str(row.iloc[0]) if len(row) > 0 and pd.notna(row.iloc[0]) else ""

            _, _, mmdd = extract_yymmdd_and_mmdd(a_val)
            if mmdd != "1001" or "10/01" in a_val or "10.01" in a_val:
                row_info = {"mmdd": mmdd}
                for item_name, col_idx in item_col_map.items():
                    cell_val = row.iloc[col_idx] if col_idx < len(row) else None
                    has_stock = (
                        pd.notna(cell_val)
                        and str(cell_val).strip() not in ["", "0", "nan", "None"]
                    )
                    row_info[item_name] = has_stock

                stock_records.append(row_info)

        return stock_records
    except Exception as e:
        st.warning(f"재고 시트 해석 중 알림: {e}")
        return []


def parse_csv_purchase(uploaded_file):
    """CSV 또는 Excel 매입 입고자료 안전하게 읽기"""
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()

    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file)
        except Exception as e:
            st.warning(f"매입 엑셀 파일 읽기 실패: {e}")
            return pd.DataFrame()

    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding=enc)
            return df
        except Exception:
            continue

    st.warning("매입 CSV 인코딩을 해석할 수 없습니다.")
    return pd.DataFrame()


def guess_purchase_code(item, prod_date, stock_records, df_csv):
    """생산일자 기준 매입자료 코드 추측"""
    lookup_item = "프릴아이스" if "프릴" in str(item) else str(item).strip()
    p_code = prefix_map.get(lookup_item, "X")

    yy_prefix, p_yymmdd, p_mmdd = extract_yymmdd_and_mmdd(prod_date)
    v_code = "K01"

    # [케이스 1] 양상추, 양배추, 적채 -> 첫 시트 재고 참조
    if lookup_item in item_col_map:
        matched_mmdd = None
        if stock_records:
            # 생산일자(p_mmdd) 이하인 재고 입고일 중 가장 최근(마지막) 날짜 선택
            valid_dates = [
                r["mmdd"]
                for r in stock_records
                if r.get(lookup_item) and r["mmdd"] <= p_mmdd
            ]
            if valid_dates:
                matched_mmdd = valid_dates[-1]

        if not matched_mmdd:
            matched_mmdd = p_mmdd

        return f"{p_code}{yy_prefix}{matched_mmdd}-{v_code}"

    # [케이스 2] 프릴 / 프릴아이스 및 기타 -> CSV/Excel 매입 입고자료에서 최신 입고분 탐색
    else:
        if isinstance(df_csv, pd.DataFrame) and not df_csv.empty:
            try:
                df_filtered = df_csv[
                    df_csv.astype(str)
                    .apply(
                        lambda r: r.str.contains(lookup_item, regex=False).any(),
                        axis=1,
                    )
                ]

                if not df_filtered.empty:
                    # 생산일자 이전 또는 당일의 최신 매입 내역 검색
                    candidate_rows = []
                    for _, r in df_filtered.iterrows():
                        row_str = " ".join(r.astype(str).tolist())
                        _, r_yymmdd, r_mmdd = extract_yymmdd_and_mmdd(row_str)
                        if r_mmdd <= p_mmdd:
                            candidate_rows.append((r_yymmdd, r_mmdd, row_str))

                    if candidate_rows:
                        # 가장 최신 입고일 선택
                        candidate_rows.sort(key=lambda x: x[0])
                        matched_yymmdd = candidate_rows[-1][0]
                        row_str = candidate_rows[-1][2]

                        vendor_find = re.search(r"\b([A-Z]\d{2})\b", row_str)
                        if vendor_find:
                            v_code = vendor_find.group(1)

                        return f"{p_code}{matched_yymmdd}-{v_code}"
            except Exception:
                pass

        return f"{p_code}{p_yymmdd}-{v_code}"


def generate_salad_report(excel_file, csv_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 첫 번째 시트 (재고 데이터) 파싱
    stock_records = parse_stock_sheet(xls)

    # 2. CSV 매입자료 파싱
    df_csv = parse_csv_purchase(csv_file)

    # 3. 일자별 생산 시트 순회 (첫 번째 시트 제외)
    exclude_sheets = ["테스트", "원가", "Sheet1"]
    daily_sheets = [
        s
        for s in xls.sheet_names[1:]
        if s not in exclude_sheets and not s.startswith("재고")
    ]

    records = []
    for sheet in daily_sheets:
        try:
            df_s = pd.read_excel(excel_file, sheet_name=sheet)

            if df_s.shape[0] < 5 or df_s.shape[1] < 14:
                continue

            prod_date = df_s.iloc[0, 0]

            items = [
                df_s.iloc[0, 4],
                df_s.iloc[0, 7],
                df_s.iloc[0, 10],
                df_s.iloc[0, 13],
            ]
            lots = [
                df_s.iloc[1, 5],
                df_s.iloc[1, 8],
                df_s.iloc[1, 11],
                df_s.iloc[1, 14],
            ]
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
                if (
                    pd.isna(item)
                    or str(item).strip() == ""
                    or str(item).strip() == "nan"
                ):
                    continue

                try:
                    p_val = float(prep) if pd.notna(prep) else 0.0
                except (ValueError, TypeError):
                    p_val = 0.0

                try:
                    a_val = float(act) if pd.notna(act) else 0.0
                except (ValueError, TypeError):
                    a_val = 0.0

                lot_str = str(lot).strip() if pd.notna(lot) else ""
                has_lot = (
                    lot_str != ""
                    and lot_str.lower() != "nan"
                    and lot_str != "none"
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
        except Exception:
            continue

    if not records:
        st.error("생산일지 시트에서 처리 가능한 데이터를 찾지 못했습니다.")
        return pd.DataFrame()

    df_report = pd.DataFrame(records)

    def process_codes(row):
        item = row["품목"]
        lot = row["작업일지_롯트"]
        has_lot = row["has_lot"]
        prod_date = row["생산일자"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        if has_lot:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = (
                f"{p}{yymmdd_worklog}-K01" if yymmdd_worklog else "없음"
            )
            code_purchase = guess_purchase_code(
                item, prod_date, stock_records, df_csv
            )
        else:
            code_worklog = "없음"
            code_purchase = guess_purchase_code(
                item, prod_date, stock_records, df_csv
            )

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
        df_result = generate_salad_report(excel_file, csv_file)
        if not df_result.empty:
            st.success("✅ 생산일지 재고 시트 및 CSV 최신 입고자료 대조가 완료되었습니다!")

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
        st.code(traceback.format_exc())
else:
    st.info("👆 1. 샐러드 생산일지 엑셀 파일과 2. CSV/Excel 매입 자료를 업로드해주세요.")
