import base64
import io
import re
import traceback
import numpy as np
import pandas as pd
import streamlit as st
from PIL import Image

# EasyOCR 호출 (이미지 분석용)
try:
    import easyocr

    reader = easyocr.Reader(["ko", "en"], gpu=False)
except Exception:
    reader = None

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

col1, col2, col3 = st.columns(3)
with col1:
    excel_file = st.file_uploader(
        "1. 샐러드 생산일지 (Excel) 업로드", type=["xlsx", "xls"]
    )
with col2:
    img_file = st.file_uploader(
        "2. (선택) 재고 표 이미지 업로드 (첫 시트 사진)",
        type=["jpg", "jpeg", "png"],
    )
with col3:
    csv_file = st.file_uploader(
        "3. 원료 매입 입고자료/수불부 (CSV/Excel)", type=["csv", "xlsx", "xls"]
    )

prefix_map = {
    "양상추": "X",
    "양배추": "D",
    "적채": "A",
    "프릴": "F",
    "프릴아이스": "F",
}

item_col_map = {
    "양상추": {"stock": 1, "input": 2},
    "양배추": {"stock": 3, "input": 4},
    "적채": {"stock": 5, "input": 6},
}


def extract_yymmdd_and_mmdd(val_str):
    """문자열/날짜 셀에서 YY, YYMMDD, MMDD 추출"""
    val_str = str(val_str).strip()

    m_full = re.search(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", val_str)
    if m_full:
        yy = m_full.group(1)[2:]
        mm = m_full.group(2).zfill(2)
        dd = m_full.group(3).zfill(2)
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_short = re.search(r"(\d{1,2})[-./](\d{1,2})", val_str)
    if m_short:
        p1 = m_short.group(1).zfill(2)
        p2 = m_short.group(2).zfill(2)
        if len(p1) == 2 and int(p1) > 12:
            yy = p1
            mm = p2
            dd = "01"
        else:
            yy = "26"
            mm = p1
            dd = p2
        return yy, f"{yy}{mm}{dd}", f"{mm}{dd}"

    m_digits = re.search(r"\b(\d{6})\b", val_str)
    if m_digits:
        d = m_digits.group(1)
        return d[:2], d, d[2:]

    return "26", "261001", "1001"


def parse_stock_image(uploaded_img):
    """업로드된 재고 표 이미지(JPG/PNG)를 EasyOCR로 읽어서 FIFO 로트 생성"""
    item_lots = {"양상추": [], "양배추": [], "적채": []}
    if reader is None:
        st.warning(
            "EasyOCR 패키지가 설치되지 않아 이미지 분석을 진행할 수 없습니다. (`pip install easyocr` 필요)"
        )
        return item_lots

    try:
        image = Image.open(uploaded_img)
        img_np = np.array(image)

        # OCR 텍스트 추출
        results = reader.readtext(img_np)
        lines = [res[1] for res in results]

        # 이미지에서 날짜 및 입고/수량 추출
        for idx, text in enumerate(lines):
            text_str = text.strip()
            if "입고" in text_str or re.search(
                r"\d{4}-\d{2}-\d{2}", text_str
            ):
                yy, yymmdd, mmdd = extract_yymmdd_and_mmdd(text_str)

                # 주변 텍스트에서 숫자(수량) 추출
                nums = []
                for offset in range(1, 8):
                    if idx + offset < len(lines):
                        clean_num = re.sub(r"[^\d]", "", lines[idx + offset])
                        if len(clean_num) >= 5:  # 10,000 이상의 수량 단위
                            nums.append(float(clean_num))

                if len(nums) >= 3:
                    item_lots["양상추"].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(nums[0] / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )
                    item_lots["양배추"].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(nums[1] / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )
                    item_lots["적채"].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(nums[2] / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )

        return item_lots
    except Exception as e:
        st.warning(f"이미지 표 파싱 중 오류 발생: {e}")
        return item_lots


def parse_stock_sheet_fifo(xls):
    """엑셀 파일 첫 번째 시트에서 FIFO 로트 파싱"""
    item_lots = {item: [] for item in item_col_map}
    try:
        stock_sheet = xls.sheet_names[0]
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
                    s_num = (
                        float(re.sub(r"[^\d.]", "", str(s_val)))
                        if s_val is not None
                        else 0.0
                    )
                except Exception:
                    s_num = 0.0

                if is_incoming and s_num > 0:
                    item_lots[item].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(s_num / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )
                elif idx == 0 and s_num > 0:
                    item_lots[item].append(
                        {
                            "yymmdd": yymmdd,
                            "avail_kg": round(s_num / 1000.0, 1),
                            "vendor": "K01",
                        }
                    )

        return item_lots
    except Exception as e:
        st.warning(f"엑셀 첫 시트 파싱 중 알림: {e}")
        return item_lots


def parse_csv_purchase(uploaded_file):
    if uploaded_file is None:
        return pd.DataFrame()

    file_name = uploaded_file.name.lower()
    if file_name.endswith(".xlsx") or file_name.endswith(".xls"):
        try:
            return pd.read_excel(uploaded_file)
        except Exception as e:
            return pd.DataFrame()

    encodings = ["utf-8-sig", "cp949", "euc-kr", "utf-8", "latin1"]
    for enc in encodings:
        try:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, encoding=enc)
        except Exception:
            continue

    return pd.DataFrame()


def get_latest_purchase_code_for_frill(item_name, prod_date, df_csv):
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, "F")
    yy_prefix, p_yymmdd, p_mmdd = extract_yymmdd_and_mmdd(prod_date)
    v_code = "F01"

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
                    row_str = " ".join(r.astype(str).tolist())
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


def consume_fifo_and_format(
    item_name, req_kg, prod_date, item_lots, df_csv
):
    lookup_item = "프릴아이스" if "프릴" in str(item_name) else str(item_name).strip()
    p_code = prefix_map.get(lookup_item, prefix_map.get(item_name, "X"))
    _, p_yymmdd, _ = extract_yymmdd_and_mmdd(prod_date)

    if lookup_item not in item_col_map:
        latest_code = get_latest_purchase_code_for_frill(
            lookup_item, prod_date, df_csv
        )
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


def generate_salad_report(excel_file, img_file, csv_file):
    xls = pd.ExcelFile(excel_file)

    # 1. 이미지가 업로드된 경우 이미지 표 OCR 우선 사용, 없을 경우 엑셀 첫 시트 파싱
    if img_file is not None:
        st.info("📷 업로드된 재고 표 이미지에서 데이터를 추출합니다.")
        item_lots = parse_stock_image(img_file)
    else:
        item_lots = parse_stock_sheet_fifo(xls)

    # 2. 매입 입고자료/수불부 파싱
    df_csv = parse_csv_purchase(csv_file)

    # 3. 일자별 생산 시트 순회
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

        if has_lot:
            _, yymmdd_worklog, _ = extract_yymmdd_and_mmdd(lot)
            code_worklog = (
                f"{p}{yymmdd_worklog}-K01" if yymmdd_worklog else "없음"
            )
        else:
            code_worklog = "없음"

        code_purchase = consume_fifo_and_format(
            item, act_kg, prod_date, item_lots, df_csv
        )

        if code_worklog != "없음" and code_worklog in code_purchase:
            is_match = "일치"
        elif code_worklog == "없음":
            is_match = "불일치(작업일지 미기재)"
        else:
            is_match = "불일치(혼합/추측)"

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
        df_result = generate_salad_report(excel_file, img_file, csv_file)
        if not df_result.empty:
            st.success("✅ 생산일지 재고 및 수불부 대조가 완료되었습니다!")

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
        st.code(traceback.format_exc())
else:
    st.info(
        "👆 1. 샐러드 생산일지 엑셀 파일과 2. (선택) 재고 표 이미지, 3. 원료 매입 입고자료/수불부 CSV를 업로드해주세요."
    )
