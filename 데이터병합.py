import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("근무기록 파일에 없는 사원번호+일자 조합만 출퇴근현황에서 가져와 자동 병합합니다.")

# 파일 업로드
caps_file = st.file_uploader("1️⃣ '출퇴근현황(캡스)' 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("2️⃣ '근무 기록(근태기록)' 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:
    try:
        # 엑셀 로드
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1)
        att_df = pd.read_excel(att_file, sheet_name=0)

        # 날짜 및 사원번호 정리
        caps_df['일자'] = pd.to_datetime(caps_df['일자'], errors='coerce').dt.date
        att_df['일자'] = pd.to_datetime(att_df['일자'], errors='coerce').dt.date
        caps_df['사원번호'] = caps_df['사원번호'].astype(str).str.zfill(5)
        att_df['사원번호'] = att_df['사원번호'].astype(str).str.zfill(5)

        # 기준 키 생성 (사원번호 + 일자)
        att_keys = set(zip(att_df['사원번호'], att_df['일자']))
        caps_df['key'] = list(zip(caps_df['사원번호'], caps_df['일자']))
        caps_df = caps_df[~caps_df['key'].isin(att_keys)]  # ❗근태기록에 없는 조합만 필터링

        # 필요한 열만 정리
        columns_to_use = ['일자', '사원번호', '소속부서', '사원명', '출근시간', '퇴근시간', '근무시간(시간단위)']
        caps_df = caps_df[columns_to_use]

        # ❗ 유효한 출퇴근 기록만 가져오기
        caps_df = caps_df[
            (caps_df['출근시간'].notna()) |
            (caps_df['퇴근시간'].notna()) |
            (caps_df['근무시간(시간단위)'].notna())
        ]

        # 병합
        merged_df = pd.concat([att_df, caps_df], ignore_index=True)

        st.success(f"✅ 근무기록에 없던 출퇴근 기록 {len(caps_df)}건이 추가되었습니다.")
        st.dataframe(merged_df)

        # 다운로드용 엑셀 파일 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='보완 근태기록')
        output.seek(0)

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
