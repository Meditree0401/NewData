import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def normalize_columns(df):
    # 컬럼명에 공백이나 특수문자 있으면 제거
    df = df.copy()
    df.columns = df.columns.str.strip()
    return df

st.set_page_config(page_title="월간 출퇴근 자동 병합 시스템", layout="wide")
st.title("📋 월간 출퇴근 자동 병합 시스템")
st.markdown("""
- 근무기록에 없는 출퇴근 내역만 병합합니다.  
- 동일 사원번호에 대해 부서명이 다를 경우, **근무기록에서 가장 마지막 날짜 기준 부서**로 통일합니다.  
- 원본 포맷은 유지되며, 병합된 행만 추가됩니다.
""")

caps_file = st.file_uploader("🟦 `출퇴근현황(캡스)` 파일을 업로드하세요", type=["xlsx"])
att_file  = st.file_uploader("🟨 `근무기록` 파일을 업로드하세요",       type=["xlsx"])

if caps_file and att_file:
    try:
        # 1) 원본 로드 (사원번호는 str로 읽어서 앞자리 0 유지)
        caps_df = pd.read_excel(caps_file, sheet_name=0, header=1, dtype={'사원번호': str})
        att_df  = pd.read_excel(att_file,  sheet_name=0,             dtype={'사원번호': str})

        # 2) 컬럼명 정규화 (공백 제거)
        caps_df = normalize_columns(caps_df)
        att_df  = normalize_columns(att_df)

        # 3) 날짜 컬럼 변환 & 유효 행 필터
        caps_df = caps_df[pd.to_datetime(caps_df['일자'], errors='coerce').notna()].copy()
        caps_df['일자_dt'] = pd.to_datetime(caps_df['일자'])
        att_df['일자_dt'] = pd.to_datetime(att_df['일자'])

        # 4) 사원명 정제 (뒤에 붙은 A, B 등 제거)
        caps_df['사원명_정규화'] = caps_df['사원명'].astype(str).str.extract(r'([가-힣]+)')
        att_df['사원명_정규화'] = att_df['사원명'].astype(str).str.extract(r'([가-힣]+)')

        # 5) 전체 데이터(caps+att)에서 사원번호별 마지막 부서 추출
        all_dept = pd.concat([
            caps_df[['사원번호','소속부서','일자_dt']],
            att_df[['사원번호','소속부서','일자_dt']]
        ], ignore_index=True)
        latest_depts = (
            all_dept.sort_values('일자_dt')
                    .groupby('사원번호')['소속부서']
                    .last()
                    .to_dict()
        )

        # 6) 출퇴근현황에 최신 부서 & 정제된 사원명 덮어쓰기
        caps_df['소속부서'] = caps_df['사원번호'].map(latest_depts).fillna(caps_df['소속부서'])
        caps_df['사원명'] = caps_df['사원명_정규화']

        # 7) 비교키 생성 (일자 + 부서 + 사원번호)
        caps_df['일자_str'] = caps_df['일자_dt'].dt.strftime('%Y-%m-%d')
        att_df['일자_str'] = att_df['일자_dt'].dt.strftime('%Y-%m-%d')
        caps_df['비교키'] = caps_df['일자_str'] + "_" + caps_df['소속부서'] + "_" + caps_df['사원번호']
        att_df['비교키'] = att_df['일자_str'] + "_" + att_df['소속부서'] + "_" + att_df['사원번호']

        # 8) 시간정보 있는 행만 필터
        time_cols = ['출근시간','퇴근시간','근무시간(시간단위)']
        caps_df_time = caps_df[caps_df[time_cols].notna().any(axis=1)].copy()

        # 9) 병합 대상 추출
        new_records = caps_df_time[~caps_df_time['비교키'].isin(att_df['비교키'])].copy()

        # 10) 혹시 모를 누락 방지용 재덮어쓰기
        new_records['소속부서'] = new_records['사원번호'].map(latest_depts).fillna(new_records['소속부서'])
        new_records['사원명']  = new_records['사원명_정규화']

        # 11) 병합용 컬럼 정리
        columns = ['일자','사원번호','소속부서','사원명',
                   '출근시간','퇴근시간','근무시간(시간단위)','근태내역','적요']
        for col in ['근태내역','적요']:
            if col not in att_df.columns:
                att_df[col] = ""
            if col not in new_records.columns:
                new_records[col] = ""

        # 12) 실제 병합
        merged_df = pd.concat([att_df[columns], new_records[columns]], ignore_index=True)

        # 13) 엑셀 생성 & 다운로드
        output = io.BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "출 퇴근현황(ID)"

        for row in dataframe_to_rows(merged_df, index=False, header=True):
            ws.append(row)

        wb.save(output)
        output.seek(0)

        st.success("✅ 병합 완료! 아래에서 최신 부서 기준 병합본을 다운로드하세요.")
        st.download_button(
            label="📥 병합된 근무기록 다운로드",
            data=output,
            file_name="근무기록_병합본_최신부서기준.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ 오류 발생: {e}")
