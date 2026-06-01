import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


TARGET_SHEET_KEYWORD = "운행기록부"
START_ROW = 15
END_ROW = 61


def to_int(value):
    if pd.isna(value) or value == "":
        return None
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    return int(float(text))


def find_period(df):
    text = " ".join(str(x) for x in df.values.flatten() if not pd.isna(x))
    m = re.search(r"(\d{4}[./-]\d{2}[./-]\d{2})\s*~\s*(\d{4}[./-]\d{2}[./-]\d{2})", text)
    if not m:
        return None, None
    start = m.group(1).replace("-", ".").replace("/", ".")
    end = m.group(2).replace("-", ".").replace("/", ".")
    return start, end


def extract_records(source_file):
    df = pd.read_excel(source_file, header=None, dtype=object)
    period_start, period_end = find_period(df)

    records = []
    date_pattern = re.compile(r"^\d{4}-\d{2}-\d{2}\([월화수목금토일]\)")

    for _, row in df.iterrows():
        first = row.iloc[0]
        if pd.isna(first):
            continue

        date_text = str(first).strip()
        if not date_pattern.match(date_text):
            continue

        weekday_match = re.search(r"\(([월화수목금토일])\)", date_text)
        weekday = weekday_match.group(1) if weekday_match else ""

        dept = row.iloc[1] if len(row) > 1 and not pd.isna(row.iloc[1]) else ""
        name = row.iloc[2] if len(row) > 2 and not pd.isna(row.iloc[2]) else ""
        start_km = to_int(row.iloc[3]) if len(row) > 3 else None
        end_km = to_int(row.iloc[4]) if len(row) > 4 else None
        distance = to_int(row.iloc[5]) if len(row) > 5 else None
        commute = to_int(row.iloc[6]) if len(row) > 6 else None
        business = to_int(row.iloc[7]) if len(row) > 7 else None

        purpose = row.iloc[8] if len(row) > 8 and not pd.isna(row.iloc[8]) else ""
        note = row.iloc[9] if len(row) > 9 and not pd.isna(row.iloc[9]) else ""
        final_note = note if str(note).strip() else purpose

        records.append({
            "date": date_text,
            "weekday": weekday,
            "dept": dept,
            "name": name,
            "start_km": start_km,
            "end_km": end_km,
            "distance": distance,
            "commute": commute,
            "business": business,
            "note": final_note,
        })

    return period_start, period_end, records


def get_target_sheet(wb):
    for ws in wb.worksheets:
        if TARGET_SHEET_KEYWORD in ws.title:
            return ws
    return wb.active


def update_template(template_file, source_file):
    period_start, period_end, records = extract_records(source_file)

    wb = load_workbook(template_file)
    ws = get_target_sheet(wb)

    if period_start:
        ws["E2"] = period_start
    if period_end:
        ws["E5"] = period_end

    # Clear old detail rows
    target_cols = ["A", "D", "E", "H", "K", "O", "S", "W", "AA", "AE"]
    for row_num in range(START_ROW, END_ROW + 1):
        for col in target_cols:
            ws[f"{col}{row_num}"] = None

    # Write new records
    for i, rec in enumerate(records):
        row_num = START_ROW + i
        if row_num > END_ROW:
            break

        ws[f"A{row_num}"] = rec["date"]
        ws[f"D{row_num}"] = rec["weekday"]
        ws[f"E{row_num}"] = rec["dept"]
        ws[f"H{row_num}"] = rec["name"]
        ws[f"K{row_num}"] = rec["start_km"]
        ws[f"O{row_num}"] = rec["end_km"]
        ws[f"S{row_num}"] = rec["distance"]
        ws[f"W{row_num}"] = rec["commute"]
        ws[f"AA{row_num}"] = rec["business"]
        ws[f"AE{row_num}"] = rec["note"]

    # Keep blank rows safe for totals
    for row_num in range(START_ROW + len(records), END_ROW + 1):
        ws[f"S{row_num}"] = 0
        ws[f"AA{row_num}"] = 0

    ws["K63"] = "=SUM(S15:V61)"
    ws["W63"] = "=SUM(W15:AA61)"
    ws["AE63"] = "=IFERROR(W63/K63,0)"

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output, records, period_start, period_end


st.title("업무용승용차 운행기록부 자동 업데이트")

st.write("제출해야 하는 운행기록부 양식과 다운로드 받은 운행내역 파일을 업로드하면 자동으로 내용을 채웁니다.")

template_file = st.file_uploader("1. 제출 양식 파일 업로드 (.xlsx)", type=["xlsx"])
source_file = st.file_uploader("2. 다운로드 받은 운행내역 파일 업로드 (.xls 또는 .xlsx)", type=["xls", "xlsx"])

if template_file and source_file:
    if st.button("운행기록부 업데이트"):
        try:
            output, records, period_start, period_end = update_template(template_file, source_file)

            st.success(f"업데이트 완료: {len(records)}건")
            if period_start and period_end:
                st.caption(f"과세기간: {period_start} ~ {period_end}")

            preview_df = pd.DataFrame(records)
            st.dataframe(preview_df, use_container_width=True)

            st.download_button(
                label="업데이트된 엑셀 다운로드",
                data=output,
                file_name="업무용운행기록부_업데이트.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error("처리 중 오류가 발생했습니다.")
            st.exception(e)
else:
    st.info("먼저 제출 양식과 운행내역 파일을 모두 업로드해 주세요.")
