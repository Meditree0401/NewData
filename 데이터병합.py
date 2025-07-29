import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")

caps_file = st.file_uploader("🟦 `출퇴근현황(캡스)` 파일을 업로드하세요", type=["xlsx"])
att_file = st.file_uploader("🟨 `근무기록` 파일을 업로드하세요", type=["xlsx"])

if caps_file and att_file:
    try:
        # 1️⃣ 파일 로딩
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1, dtype={'사원번호': str})
        att_df = pd.read_excel(att_file, sheet_name=0, dtype={'사원번호': str})

        # → 컬럼명 확인
        st.write("▶ caps_df columns:", caps_df.columns.tolist())
        st.write("▶ att_df columns:", att_df.columns.tolist())

        # 2️⃣ 날짜 정리
        caps_df = caps_df[pd.to_datetime(caps_df['일자'], errors='coerce').notna()].copy()
        caps_df['일자_dt'] = pd.to_datetime(caps_df['일자'])
        att_df['일자_dt'] = pd.to_datetime(att_df['일자'])

        # 3️⃣ 사원명 정규화
        caps_df['사원명_정규화'] = caps_df['사원명'].str.extract(r'([가-힣]+)')
        att_df['사원명_정규화'] = att_df['사원명'].str.extract(r'([가-힣]+)')

        # 4️⃣ 최신 부서 매핑
        all_df = pd.concat([
            caps_df[['사원번호', '소속부서', '일자_dt']],
            att_df[['사원번호', '소속부서', '일자_dt']]
        ], ignore_index=True)
        latest_depts = (all_df
                        .sort_values('일자_dt')
                        .groupby('사원번호')['소속부서']
                        .last()
                        .to_dict())
        st.write("▶ latest_depts sample:", dict(list(latest_depts.items())[:5]))

        # 5️⃣ 부서 및 이름 덮어쓰기
        caps_df['소속부서'] = caps_df['사원번호'].map(latest_depts).fillna(caps_df['소속부서'])
        caps_df['사원명'] = caps_df['사원명_정규화']

        # 6️⃣ 비교키 생성
        caps_df['일자_str'] = caps_df['일자_dt'].dt.strftime('%Y-%m-%d')
        att_df['일자_str'] = att_df['일자_dt'].dt.strftime('%Y-%m-%d')
        caps_df['비교키'] = caps_df['일자_str'] + "_" + caps_df['소속부서'] + "_" + caps_df['사원번호']
        att_df['비교키'] = att_df['일자_str'] + "_" + att_df['소속부서'] + "_" + att_df['사원번호']

        # 7️⃣ 시간 정보 있는 행만 필터
        time_cols = ['출근시간', '퇴근시간', '근무시간(시간단위)']
        caps_df_time = caps_df[caps_df[time_cols].notna().any(axis=1)].copy()
        st.write("▶ caps_df_time sample:", caps_df_time.head())

        # 8️⃣ 병합 대상 추출
        new_records = caps_df_time[~caps_df_time['비교키'].isin(att_df['비교키'])].copy()
        st.write("▶ new_records sample:", new_records[['사원번호','소속부서']].drop_duplicates().head())

        # (이제 디버그 확인 후, 아래부터는 기존 병합 로직을 그대로 이어가시면 됩니다)
        # …
        # st.download_button(…) 으로 다운로드 버튼 표시

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
