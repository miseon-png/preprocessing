import base64
import io
import re
import traceback
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="샐러드 생산일보 원료코드 대조", layout="wide")

# 1. 인쇄 및 공통 스타일 설정
st.markdown(
    """
    <style>
    @media print {
        header, footer, .stFileUploader, button, iframe, [data-testid="stHeader"], .stAlert, .no-print, [data-testid="stToolbar"] {
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
        }
        .print-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 10px;
            background-color: #ffffff !important;
            color: #000000 !important;
        }
        .print-table th, .print-table td {
            border: 1px solid #333333 !important;
            padding: 5px 4px;
            text-align: center;
            white-space: nowrap;
        }
        .print-table th {
            background-color: #f2f2f2 !important;
            font-weight: bold;
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
        "1. 샐러드 생산일지 (Excel) 업로드 (첫 시트 또는 Sheet1: 당월 재고)",
        type=["xlsx", "xls"],
    )
with col2:
    csv_file = st.file_uploader(
        "2. 원료 매입 입고자료 / 수불부 (CSV/Excel) 업로드",
        type=["csv", "xlsx", "xls"],
    )

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}

# 재고 시트 컬럼 인덱스 매핑 (A=날짜/입고, B=양상추재고, C=양상추투입, D=양배추재고...)
item_col_map = {
    "양상추": {"stock": 1, "input": 2},
    "양배추": {"stock": 3, "input": 4},
    "적채": {"stock": 5, "input": 6},
}


def extract_yymmdd_and_mmdd(val_str):
    """문자열/날짜 셀에서 연, 월, 일 추출"""
    val_str = str(val_str).strip()
    current_yy = datetime.now().strftime("%y")  # 현재 연도 2자리 (예: '26')

    # 1. YYYY-MM-DD 또는 YYYY.MM.DD
    m_full = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        yy = m_full.group(1)[2:]
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    # 2. MM/DD 또는 MM.DD (입고 문구 등)
    m_short = re.search(r"(\d{1,2})[-./](\d{1,2})", val_str)
    if m_short:
        p1 = m_short.group(1).zfill(2)
        p2 = m_short.group(2).zfill(2)
        if len(p1) == 2 and int(p1) > 12:
            yy = p1
            mm = p2
            dd = "01"
        else:
            yy = current_yy
            mm = p1
            dd = p2
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    # 3. YYMMDD 6자리
    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]

    # 기본값
    today_yymmdd = datetime.now().strftime("%y%m%d")
    return current_yy, today_yymmdd, today_yymmdd[2:]


def find_stock_sheet_name(xls):
    """
    재고 시트 자동 탐색 로직:
    1. 'sheet1' (대소문자 구분 없음) 시트가 존재하면 해당 시트 사용
    2. 없으면 시트들을 순회하며 '양상추', '양배추', '적채' 키워드가 들어있는 헤더를 가진 시트 탐색
    3. 모두 없을 경우 첫 번째 시트 반환
    """
    for sheet_name in xls.sheet_names:
        if sheet_name.strip().lower() == "sheet1":
            return sheet_name

    keywords = ["양상추", "양배추", "적채"]
    for sheet_name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None, nrows=10)
            for idx, row in df.iterrows():
                row_str = " ".join(row.dropna().astype(str).tolist())
                if all(k in row_str for k in keywords):
                    return sheet_name
        except Exception:
            continue

    return xls.sheet_names[0]


def parse_stock_sheet_fifo(xls):
    """재고 시트에서 양상추, 양배추, 적채의 초기 재고 및 입고 로트 큐 생성"""
    item_lots = {item: [] for item in item_col_map}
    try:
        stock_sheet = find_stock_sheet_name(xls)
        df_stock = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

        for idx, row in df_stock.iterrows():
            a_val = (
                str(row.iloc[0]).strip()
                if len(row) > 0 and pd.notna(row.iloc[0])
                else ""
            )
            if not a_val or a_val.lower() == "nan":
                continue

            yy, yymmdd, mmdd = extract_yymmdd_and_mmdd(a_val)
            is_incoming = "입고" in a_val

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

                if (is_incoming or idx == 0) and s_num > 0:
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


def parse_csv_purchase(uploaded_file):
    """원료 매입 입고자료 / 수불부 파일 파싱"""
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file)
        except Exception:
            return pd.DataFrame()

    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except Exception:
            continue

    return pd.DataFrame()


def get_latest_purchase_code(item_name, prod_date, df_csv):
    """수불부(CSV/Excel)에서 생산일자 이전/당일 기준 해당 품목의 가장 최근 매입/입고 로트 코드 조회"""
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, prefix_map.get(str(item_name).strip(), "X"))
    yy_prefix, p_yymmdd, p_mmdd = extract_yymmdd_and_mmdd(prod_date)
    v_code = "F01" if p_code == "F" else "K01"

    if isinstance(df_csv, pd.DataFrame) and not df_csv.empty:
        try:
            df_filtered = df_csv[
                df_csv.astype(str).apply(
                    lambda r: r.str.contains(lookup_item, regex=False).any(),
                    axis=1,
                )
            ]

            if not df_filtered.empty:
                candidate_rows = []
                for _, r in df_filtered.iterrows():
                    row_str = " ".join(r.dropna().astype(str).tolist())
                    _, r_yymmdd, r_mmdd = extract_yymmdd_and_mmdd(row_str)
                    if r_mmdd <= p_mmdd:
                        candidate_rows.append((r_yymmdd, r_mmdd, row_str))

                if candidate_rows:
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


def consume_fifo_and_format(item_name, req_kg, prod_date, item_lots, df_csv):
    """
    1) 양상추/양배추/적채: 첫 시트/Sheet1 재고 기반 FIFO 소진 + 수량 병기
    2) 프릴/프릴아이스: 수불부(CSV) 기준 최근 입고일자(yymmdd) 매칭
    """
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, prefix_map.get(item_name, "X"))
    _, p_yymmdd, _ = extract_yymmdd_and_mmdd(prod_date)

    if lookup_item not in item_col_map:
        latest_code = get_latest_purchase_code(lookup_item, prod_date, df_csv)
        return f"{latest_code}({round(req_kg, 1)}kg)"

    lots = item_lots.get(lookup_item, [])

    if not lots or req_kg <= 0:
        return f"{p_code}{p_yymmdd}-K01({round(req_kg, 1)}kg)"

    remaining = req_kg
    used_parts = []

    for lot in lots:
        if remaining <= 0:
            break
        if lot["avail_kg"] <= 0:
            continue

        used = min(lot["avail_kg"], remaining)
        lot["avail_kg"] -= used
        remaining -= used

        code = f"{p_code}{lot['yymmdd']}-{lot.get('vendor', 'K01')}"
        used_parts.append(f"{code}({round(used, 1)}kg)")

    if remaining > 0:
        latest_yymmdd = lots[-1]["yymmdd"] if lots else p_yymmdd
        v_code = lots[-1].get("vendor", "K01") if lots else "K01"
        code = f"{p_code}{latest_yymmdd}-{v_code}"
        used_parts.append(f"{code}({round(remaining, 1)}kg)")

    return " / ".join(used_parts)


def generate_salad_report(excel_file, csv_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 재고 시트 자동 감지 및 파싱
    item_lots, stock_sheet_name = parse_stock_sheet_fifo(xls)

    # 2. 원료 매입 입고자료 / 수불부 파싱
    df_csv = parse_csv_purchase(csv_file)

    # 3. 일자별 생산 시트 순회 (작업일지가 아닌 시트 제외)
    exclude_sheets = ["테스트", "원가", stock_sheet_name, "Sheet1", "sheet1"]
    daily_sheets = [
        s
        for s in xls.sheet_names
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
                    or str(item).strip().lower() == "nan"
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
                    and lot_str.lower() != "none"
                )

                records.append(
                    {
                        "생산일자": prod_date,
                        "품목": str(item).strip(),
                        "작업일지_롯트": lot_str if has_lot else "없음",
                        "실투입 양 (kg)": round(a_val / 1000.0, 1),
                        "준비 양 (kg)": round(p_val / 1000.0, 1),
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
        act_kg = row["실투입 양 (kg)"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))

        # 1. 작업일지 기준 코드
        if has_lot:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = (
                f"{p}{yymmdd_worklog}-K01" if yymmdd_worklog else "없음"
            )
        else:
            code_worklog = "없음"

        # 2. 수불부 기준 코드 (무조건 가장 최근 매입/입고 내역 조회)
        code_subul = get_latest_purchase_code(item, prod_date, df_csv)

        # 3. 매입자료 기준 코드 (재고 시트 FIFO 소진)
        code_purchase = consume_fifo_and_format(
            item, act_kg, prod_date, item_lots, df_csv
        )

        # 4. 일치 여부
        if code_worklog != "없음" and code_worklog in code_purchase:
            is_match = "일치"
        elif code_worklog == "없음":
            is_match = "불일치(작업일지 미기재)"
        else:
            is_match = "불일치(혼합/추측)"

        return pd.Series([code_worklog, code_subul, code_purchase, is_match])

    df_report[["작업일지 기준 코드", "수불부 기준 코드", "매입자료 기준 코드", "일치여부"]] = (
        df_report.apply(process_codes, axis=1)
    )

    return df_report[
        [
            "생산일자",
            "품목",
            "작업일지 기준 코드",
            "수불부 기준 코드",
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
            st.success("✅ 원료코드 대조 작업이 완료되었습니다!")

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

            st.download_button(
                label="📥 엑셀 파일 다운로드 (.xlsx)",
                data=buffer.getvalue(),
                file_name="원료코드_비교_정리표.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            st.caption("💡 Tip: 인쇄를 원하시면 브라우저 인쇄 단축키(`Ctrl+P` 또는 `Cmd+P`)를 눌러주세요.")

    except Exception as e:
        st.error(f"처리 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    st.info(
        "👆 1. 샐러드 생산일지 엑셀 파일과 2. 원료 매입 입고자료/수불부 CSV/Excel을 업로드해주세요."
    )
