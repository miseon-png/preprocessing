import base64
import copy
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
        "1. 샐러드 생산일지 (Excel) 업로드 (첫 시트 또는 Sheet1: 당월 재고/매입자료)",
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

# B=1: 양상추 매입, C=2: 양상추 재고 / E=4: 양배추 매입, F=5: 양배추 재고 / H=7: 적채 매입, I=8: 적채 재고
item_col_map = {
    "양상추": {"purchase_col": 1, "stock_col": 2},
    "양배추": {"purchase_col": 4, "stock_col": 5},
    "적채": {"purchase_col": 7, "stock_col": 8},
}


def extract_yymmdd_and_mmdd(val):
    if pd.isna(val):
        today = datetime.now()
        return today.strftime("%y"), today.strftime("%y%m%d"), today.strftime("%m%d")

    if isinstance(val, (pd.Timestamp, datetime)):
        yy = val.strftime("%y")
        yymmdd = val.strftime("%y%m%d")
        mmdd = val.strftime("%m%d")
        return yy, yymmdd, mmdd

    val_str = str(val).strip()

    m_full = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        yy = m_full.group(1)[2:]
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    try:
        dt = pd.to_datetime(val_str, errors="coerce", yearfirst=True)
        if pd.notna(dt):
            yy = dt.strftime("%y")
            yymmdd = dt.strftime("%y%m%d")
            mmdd = dt.strftime("%m%d")
            return yy, yymmdd, mmdd
    except Exception:
        pass

    current_yy = datetime.now().strftime("%y")

    m_short = re.search(r"(\d{1,2})[-./](\d{1,2})", val_str)
    if m_short:
        p1 = m_short.group(1).zfill(2)
        p2 = m_short.group(2).zfill(2)
        if int(p1) > 12:
            yy = p1
            mm = p2
            dd = "01"
        else:
            yy = current_yy
            mm = p1
            dd = p2
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]

    today = datetime.now()
    return current_yy, today.strftime("%y%m%d"), today.strftime("%m%d")


def find_stock_sheet_name(xls):
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
    """
    매입자료 시트 파싱:
    - B, E, H열 매입 수량이 있는 행만 매입일자 로트로 등록
    - 첫 번째 유효 데이터 행의 재고(C, F, I열)는 기초재고 로트로 등록
    """
    item_lots = {item: [] for item in item_col_map}
    try:
        stock_sheet = find_stock_sheet_name(xls)
        df_stock = pd.read_excel(xls, sheet_name=stock_sheet, header=None)

        found_first_data_row = False

        for idx, row in df_stock.iterrows():
            if len(row) == 0 or pd.isna(row.iloc[0]):
                continue

            raw_date = row.iloc[0]
            a_val = str(raw_date).strip()
            if not a_val or a_val.lower() == "nan" or "날짜" in a_val:
                continue

            yy, yymmdd, mmdd = extract_yymmdd_and_mmdd(raw_date)

            is_first_data_row = False
            if not found_first_data_row and re.search(r"\d", a_val):
                is_first_data_row = True
                found_first_data_row = True

            for item, cols in item_col_map.items():
                p_col = cols["purchase_col"]
                s_col = cols["stock_col"]

                p_val = (
                    row.iloc[p_col]
                    if p_col < len(row) and pd.notna(row.iloc[p_col])
                    else None
                )
                try:
                    cleaned_p = re.sub(r"[^\d.]", "", str(p_val))
                    p_num = float(cleaned_p) if cleaned_p else 0.0
                except Exception:
                    p_num = 0.0

                if p_num > 0:
                    item_lots[item].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(p_num / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )
                elif is_first_data_row:
                    s_val = (
                        row.iloc[s_col]
                        if s_col < len(row) and pd.notna(row.iloc[s_col])
                        else None
                    )
                    try:
                        cleaned_s = re.sub(r"[^\d.]", "", str(s_val))
                        s_num = float(cleaned_s) if cleaned_s else 0.0
                    except Exception:
                        s_num = 0.0

                    if s_num > 0:
                        item_lots[item].append(
                            {
                                "yymmdd": yymmdd,
                                "avail_kg": round(s_num / 1000.0, 1),
                                "vendor": "K01",
                            }
                        )

        return item_lots, stock_sheet
    except Exception as e:
        st.warning(f"매입자료 시트 파싱 알림: {e}")
        return item_lots, xls.sheet_names[0]


def parse_csv_purchase(uploaded_file):
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
    """2번 수불부(CSV) 기준 최신 매입 코드 추출"""
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, prefix_map.get(str(item_name).strip(), "X"))
    _, p_yymmdd, p_mmdd = extract_yymmdd_and_mmdd(prod_date)
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

                    if r_yymmdd <= p_yymmdd:
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


def generate_salad_report(excel_file, csv_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 매입자료 시트 원본 로트 파싱
    master_item_lots, stock_sheet_name = parse_stock_sheet_fifo(xls)
    running_item_lots = copy.deepcopy(master_item_lots)

    # 2. 원료 매입 입고자료 / 수불부 파싱
    df_csv = parse_csv_purchase(csv_file)

    # 3. 일자별 생산 시트 순회
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

    # 순서 보장을 위한 순차 루프 기반 코드 생성
    final_rows = []
    for rec in records:
        item = rec["품목"]
        lot = rec["작업일지_롯트"]
        has_lot = rec["has_lot"]
        prod_date = rec["생산일자"]
        act_kg = rec["실투입 양 (kg)"]

        lookup_item = "프릴아이스" if "프릴" in item else item
        p = prefix_map.get(lookup_item, prefix_map.get(item, "X"))
        _, p_yymmdd, _ = extract_yymmdd_and_mmdd(prod_date)

        # 1. 작업일지 기준 코드
        if has_lot:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = (
                f"{p}{yymmdd_worklog}-K01" if yymmdd_worklog else "없음"
            )
        else:
            code_worklog = "없음"

        # 2. 수불부 기준 코드 (CSV 파일 기반)
        code_subul = get_latest_purchase_code(item, prod_date, df_csv)

        # 3. 매입자료 기준 코드 (순차 FIFO 차감)
        if lookup_item not in item_col_map:
            code_purchase = ""
        else:
            all_lots = running_item_lots.get(lookup_item, [])
            valid_lots = [l for l in all_lots if l["yymmdd"] <= p_yymmdd]

            if not valid_lots or act_kg <= 0:
                code_purchase = f"{p}{p_yymmdd}-K01({round(act_kg, 1)}kg)"
            else:
                remaining = act_kg
                used_parts = []

                for lot_item in valid_lots:
                    if remaining <= 0:
                        break
                    if lot_item["avail_kg"] <= 0:
                        continue

                    used = min(lot_item["avail_kg"], remaining)
                    lot_item["avail_kg"] -= used
                    remaining -= used

                    code_str = f"{p}{lot_item['yymmdd']}-{lot_item.get('vendor', 'K01')}"
                    used_parts.append(f"{code_str}({round(used, 1)}kg)")

                if remaining > 0:
                    latest_yymmdd = valid_lots[-1]["yymmdd"] if valid_lots else p_yymmdd
                    v_code = valid_lots[-1].get("vendor", "K01") if valid_lots else "K01"
                    code_str = f"{p}{latest_yymmdd}-{v_code}"
                    used_parts.append(f"{code_str}({round(remaining, 1)}kg)[부족]")

                code_purchase = " / ".join(used_parts)

        # 4. 일치 여부 판정
        if code_worklog != "없음" and code_purchase != "" and code_worklog in code_purchase:
            is_match = "일치"
        elif code_worklog != "없음" and code_purchase == "" and code_worklog == code_subul:
            is_match = "일치"
        elif code_worklog == "없음":
            is_match = "불일치(작업일지 미기재)"
        else:
            is_match = "불일치(혼합/추측)"

        rec["작업일지 기준 코드"] = code_worklog
        rec["수불부 기준 코드"] = code_subul
        rec["매입자료 기준 코드"] = code_purchase
        rec["일치여부"] = is_match
        final_rows.append(rec)

    df_report = pd.DataFrame(final_rows)

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
