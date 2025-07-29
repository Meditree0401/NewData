import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("근무기록에 없는 출퇴근 내역만 병합하며, 사원명과 부서는 반드시 근무기록 기준으로 통일합니다.")

# 파일 업로드
caps_file = st.file_uploader("1️⃣ '출퇴근현황(캡스)' 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("2️⃣ '근무기록' 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:
    try:
        # 엑셀 로딩
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1)
        att_df = pd.read_excel(att_file, sheet_name=0)

        # 날짜 형식 정리 & 사원번호 통일 (숫자만, zfill 5자리)
        for df in [caps_df, att_df]:
            df['일자'] = pd.to_datetime(df['일자'], errors='coerce').dt.date
            df['사원번호'] = df['사원번호'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(5)

        # 근무기록 기준 사원번호 → 이름/부서 매핑 딕셔너리 생성
        id_to_name = att_df.set_index('사원번호')['사원명'].to_dict()
        id_to_dept = att_df.set_index('사원번호')['소속부서'].to_dict()

        # 출퇴근현황에 이름/부서 덮어쓰기: 무조건 근무기록 기준
        caps_df['사원번호'] = caps_df['사원번호'].astype(str).str.replace(r'\D', '', regex=True).str.zfill(5)
        caps_df['사원명'] = caps_df['사원번호'].map(id_to_name).fillna(caps_df['사원명'])
        caps_df['소속부서'] = caps_df['사원번호'].map(id_to_dept).fillna(caps_df['소속부서'])

        # 기준 키 생성 (사원번호 + 일자)
        att_keys = set(zip(att_df['사원번호'], att_df['일자']))
        caps_df['key'] = list(zip(caps_df['사원번호'], caps_df['일자']))
        new_records = caps_df[~caps_df['key'].isin(att_keys)]

        # 유효한 출퇴근 내역만 필터링
        new_records = new_records[
            (new_records['출근시간'].notna()) |
            (new_records['퇴근시간'].notna()) |
            (new_records['근무시간(시간단위)'].notna())
        ]

        # 필요한 열 정리
        columns_to_use = ['일자', '사원번호', '소속부서', '사원명', '출근시간', '퇴근시간', '근무시간(시간단위)']
        new_records = new_records[columns_to_use]

        # 병합
        merged_df = pd.concat([att_df, new_records], ignore_index=True)

        st.success(f"✅ 근무기록에 없던 출퇴근 내역 {len(new_records)}건이 추가되었습니다. (사원명·부서 통일 완료)")
        st.dataframe(merged_df)

        # 엑셀 저장
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
        st.error(f"⚠️ 오류 발생: {e}")

else:
    st.info("👆 두 파일을 모두 업로드하면 병합 결과가 표시됩니다.")
