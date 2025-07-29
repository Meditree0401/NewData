import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("근무기록 기준으로 사원번호, 사원명, 소속부서를 완전히 일치시켜 병합합니다.")

caps_file = st.file_uploader("1️⃣ '출퇴근현황(캡스)' 파일 업로드", type=["xlsx"])
att_file = st.file_uploader("2️⃣ '근무기록' 파일 업로드", type=["xlsx"])

if caps_file and att_file:
    try:
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1)
        att_df = pd.read_excel(att_file, sheet_name=0)

        # 날짜 및 사원번호 정리 (모두 근무기록 기준에 맞춤)
        for df in [caps_df, att_df]:
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce').dt.date
            df['사원번호'] = df['사원번호'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(5)

        # 사원번호 기준 이름과 부서 매핑
        id_to_name = att_df.set_index('사원번호')['사원명'].to_dict()
        id_to_dept = att_df.set_index('사원번호')['소속부서'].to_dict()

        # 출퇴근현황에 이름/부서 덮어쓰기
        caps_df['사원명'] = caps_df['사원번호'].map(id_to_name).fillna(caps_df['사원명'])
        caps_df['소속부서'] = caps_df['사원번호'].map(id_to_dept).fillna(caps_df['소속부서'])

        # 기준 키: 사원번호 + 일자
        att_keys = set(zip(att_df['사원번호'], att_df['일자']))
        caps_df['key'] = list(zip(caps_df['사원번호'], caps_df['일자']))
        caps_df = caps_df[~caps_df['key'].isin(att_keys)]

        # 필요한 열만 사용
        columns_to_use = ['일자', '사원번호', '소속부서', '사원명', '출근시간', '퇴근시간', '근무시간(시간단위)']
        caps_df = caps_df[columns_to_use]

        # 출근, 퇴근, 근무시간 중 하나라도 있는 행만
        caps_df = caps_df[
            (caps_df['출근시간'].notna()) |
            (caps_df['퇴근시간'].notna()) |
            (caps_df['근무시간(시간단위)'].notna())
        ]

        # 병합
        merged_df = pd.concat([att_df, caps_df], ignore_index=True)

        st.success(f"✅ 근무기록에 없던 출퇴근 기록 {len(caps_df)}건이 근무기록 기준으로 통일되어 추가되었습니다.")
        st.dataframe(merged_df)

        # 다운로드용 엑셀 저장
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
        st.error(f"⚠️ 오류 발생: {e}")

else:
    st.info("👆 두 파일을 모두 업로드하면 병합 결과가 나타납니다.")
