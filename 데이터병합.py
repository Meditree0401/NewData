import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import tempfile

st.title("월간 출퇴근 자동 병합 시스템")

# 1. 파일 업로드
caps_file = st.file_uploader("📥 '출퇴근현황(캡스)' 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("📥 '근무 기록(근태기록)' 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:

    # 2. 엑셀 파일 읽기
    caps_xl = pd.ExcelFile(caps_file)
    att_xl = pd.ExcelFile(att_file)

    # 캡스 파일은 두 번째 줄부터 데이터 시작
    caps_df = pd.read_excel(caps_xl, sheet_name=caps_xl.sheet_names[0], skiprows=1)
    att_df = pd.read_excel(att_xl, sheet_name=att_xl.sheet_names[0])

    # 3. 전처리
    caps_df.columns = caps_df.columns.str.strip()
    caps_df['일자'] = pd.to_datetime(caps_df['일자'], errors='coerce')
    caps_df['사원번호'] = caps_df['사원번호'].astype(str).str.zfill(5)

    att_df['일자'] = pd.to_datetime(att_df['일자'], errors='coerce')
    att_df['사원번호'] = att_df['사원번호'].astype(str).str.zfill(5)

    # 4. 기존 근무기록에 없는 (일자 + 사원번호) 찾기
    merged = pd.merge(
        caps_df,
        att_df[['일자', '사원번호']],
        on=['일자', '사원번호'],
        how='left',
        indicator=True
    )
    new_data = merged[merged['_merge'] == 'left_only']
    new_data = new_data[caps_df.columns]  # 순서 유지

    # 5. 근태기록 엑셀을 임시 파일로 저장한 뒤 openpyxl로 열기
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(att_file.read())
        tmp_path = tmp.name

    try:
        wb = load_workbook(tmp_path)
        ws = wb[wb.sheetnames[0]]
        start_row = ws.max_row + 1

        # 6. 새 데이터를 기존 시트에 추가
        for _, row in new_data.iterrows():
            for col_idx, val in enumerate(row, start=1):
                ws.cell(row=start_row, column=col_idx, value=val)
            start_row += 1

        # 7. 병합 완료된 파일 다운로드 제공
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as out_tmp:
            wb.save(out_tmp.name)
            st.success("✅ 병합이 완료되었습니다! 아래에서 파일을 다운로드하세요.")
            st.download_button(
                label="📤 병합된 파일 다운로드",
                data=open(out_tmp.name, "rb"),
                file_name="병합된_근무기록.xlsx"
            )

    except Exception as e:
        st.error(f"❌ 파일 병합 중 오류가 발생했습니다:\n\n{str(e)}")
