import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("두 개의 엑셀 파일을 업로드하면, 누락된 출퇴근 기록을 자동으로 병합하여 다운로드할 수 있습니다.")

# 파일 업로드
caps_file = st.file_uploader("1️⃣ '출퇴근현황(캡스)' 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("2️⃣ '근무 기록(근태기록)' 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:
    try:
        # 엑셀 로드
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1)  # 두 번째 줄이 컬럼
        att_df = pd.read_excel(att_file, sheet_name=0)

        # 날짜 정리
        caps_df['일자'] = pd.to_datetime(caps_df['일자'], errors='coerce').dt.date
        att_df['일자'] = pd.to_datetime(att_df['일자'], errors='coerce').dt.date

        # 사원번호 자리수 맞추기 (5자리로 통일)
        caps_df['사원번호'] = caps_df['사원번호'].astype(str).str.zfill(5)
        att_df['사원번호'] = att_df['사원번호'].astype(str).str.zfill(5)

        # 기준 키 생성 (사원번호 + 일자)
        att_keys = set(zip(att_df['사원번호'], att_df['일자']))
        caps_keys = set(zip(caps_df['사원번호'], caps_df['일자']))
        missing_keys = caps_keys - att_keys

        # 누락된 데이터만 필터링
        missing_df = caps_df[caps_df.set_index(['사원번호', '일자']).index.isin(missing_keys)]

        # 근태기록 파일에 맞는 열만 선택
        columns_to_use = ['일자', '사원번호', '소속부서', '사원명', '출근시간', '퇴근시간', '근무시간(시간단위)']
        missing_df = missing_df[columns_to_use]

        # 병합
        merged_df = pd.concat([att_df, missing_df], ignore_index=True)

        st.success(f"✅ 누락된 출퇴근 기록 {len(missing_df)}건이 추가되었습니다.")
        st.dataframe(merged_df)

        # 엑셀로 저장
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='보완 근태기록')
        output.seek(0)

        # 다운로드 버튼
        st.download_button(
            label="📥 보완된 근태기록 다운로드",
            data=output,
            file_name="보완_근태기록.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ 오류가 발생했습니다: {e}")

else:
    st.info("👆 위의 두 파일을 모두 업로드하면 자동으로 병합됩니다.")
