import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("근무기록에 없는 출퇴근 내역만 병합하며, **사원명과 부서는 반드시 근무기록 기준**으로 통일합니다.")

# 파일 업로드
caps_file = st.file_uploader("🟦 `출퇴근현황(캡스)` 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("🟨 `근무기록` 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:
    try:
        # 파일 로딩
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1)
        att_df = pd.read_excel(att_file, sheet_name=0)

        # 날짜 필터링 및 사원번호 처리
        for df in [caps_df, att_df]:
            df['사원번호'] = df['사원번호'].astype(str)

        # 날짜 형식 정리
        caps_df = caps_df[pd.to_datetime(caps_df['일자'], errors='coerce').notna()].copy()
        att_df['일자_str'] = pd.to_datetime(att_df['일자']).dt.strftime('%Y-%m-%d')
        caps_df['일자_str'] = pd.to_datetime(caps_df['일자']).dt.strftime('%Y-%m-%d')

        # 사원명에서 뒤 A 제거
        caps_df['사원명_정규화'] = caps_df['사원명'].astype(str).str.extract(r'([가-힣]+)')
        att_df['사원명_정규화'] = att_df['사원명'].astype(str).str.extract(r'([가-힣]+)')

        # 비교키 생성
        att_df['비교키'] = att_df['일자_str'] + "_" + att_df['소속부서'].astype(str) + "_" + att_df['사원명_정규화']
        caps_df['비교키'] = caps_df['일자_str'] + "_" + caps_df['소속부서'].astype(str) + "_" + caps_df['사원명_정규화']

        # 누락된 내역만 추출 + 출퇴근 시간 있는 것만
        new_records = caps_df[~caps_df['비교키'].isin(att_df['비교키'])].copy()
        new_records = new_records[
            new_records[['출근시간', '퇴근시간', '근무시간(시간단위)']].notna().any(axis=1)
        ].copy()

        # 사원명 정리
        new_records['사원명'] = new_records['사원명_정규화']

        # 필요한 열 맞추기
        columns = ['일자', '사원번호', '소속부서', '사원명', 
                   '출근시간', '퇴근시간', '근무시간(시간단위)', '근태내역', '적요']

        for col in ['근태내역', '적요']:
            att_df[col] = ""
            new_records[col] = ""

        formatted_new = new_records[columns].copy()
        original_data = att_df[columns].copy()

        # 병합
        merged_df = pd.concat([original_data, formatted_new], ignore_index=True)

        # 엑셀로 변환
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "출 퇴근현황(ID)"

        for row in dataframe_to_rows(merged_df, index=False, header=True):
            ws.append(row)

        wb.save(output)
        output.seek(0)

        st.success("✅ 병합 완료! 아래에서 파일을 다운로드하세요.")
        st.download_button(
            label="📥 병합된 근무기록 다운로드",
            data=output,
            file_name="근무기록_병합본.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
