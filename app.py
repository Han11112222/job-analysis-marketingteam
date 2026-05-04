import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import os

# 1. 페이지 설정
st.set_page_config(page_title="직무분석 마케팅팀", layout="wide")

st.title("📊 직무분석 마케팅팀 (2025 vs 2026)")
st.markdown("""
**작성자:** 최한엽 (Han)  
2025년 대비 2026년 마케팅팀의 **인원 감축(-2명)** 및 **조직 개편(마케팅 1, 2팀 → 통합 부팀)** 상황을 반영한 직무 분석 대시보드입니다.  
웹앱 자동화 및 데이터 분석을 통해 물리적 인원 감소를 극복하고 업무 효율화를 이뤄낸 성과를 시각화했습니다.
""")

# 2. 데이터 로딩 함수 (캐싱 적용)
@st.cache_data
def load_data(file_path):
    # 엑셀 파일 읽기
    df = pd.read_excel(file_path, sheet_name='직무표')
    
    # 컬럼명 단순화 및 전처리
    df = df.rename(columns={
        '직무\n(Job)': '직무',
        '책무\n(Duty)': '책무',
        '과업\n(Task)': '과업',
        '과업\n수행시간\n(연간)': '과업수행시간'
    })
    
    # 중복 컬럼 제거 (헤더 병합 시 발생하는 문제 방지)
    df = df.loc[:, ~df.columns.duplicated()]
    
    # 병합된 셀(빈칸) 데이터 앞의 값으로 채우기
    df['직무'] = df['직무'].ffill()
    df['책무'] = df['책무'].ffill()
    
    # 수행시간이 숫자인 데이터만 남기기
    df['과업수행시간'] = pd.to_numeric(df['과업수행시간'], errors='coerce')
    df = df.dropna(subset=['과업수행시간'])
    
    return df

# 3. 깃허브에 업로드된 파일명 (동일 경로)
file_2025 = "직무현황표_20250515_2025B_마케팅팀.xlsx"
file_2026 = "직무현황표_20260408_2026A_마케팅팀.xlsx"

# 파일 존재 여부 확인 후 실행
if os.path.exists(file_2025) and os.path.exists(file_2026):
    df_2025 = load_data(file_2025)
    df_2026 = load_data(file_2026)
    
    # 시간 계산 (1인당 표준 근무: 1826.7시간)
    std_hours = 1826.7
    reduced_people = 2
    loss_hours = std_hours * reduced_people # -3,653.4 시간
    
    sum_2025 = df_2025['과업수행시간'].sum()
    sum_2026 = df_2026['과업수행시간'].sum()
    
    # 조직통폐합/자동화로 추가 세이브된 시간 계산
    efficiency_hours = sum_2025 - loss_hours - sum_2026
    
    st.divider()
    
    # =====================================================================
    # 시각화 1: 워터폴 차트 (인원 감축 방어 및 효율화)
    # =====================================================================
    st.subheader("💡 1. 인원 감소 방어 및 자동화 업무 효율화 성과")
    st.markdown(f"총 **{loss_hours:,.1f}시간(2명)**의 T/O 감축이 있었으나, 부팀 통폐합 및 파이썬 웹앱 도입을 통해 오히려 **{efficiency_hours:,.1f}시간**의 추가 비효율을 걷어냈습니다.")

    fig_waterfall = go.Figure(go.Waterfall(
        name = "Hours",
        orientation = "v",
        measure = ["absolute", "relative", "relative", "total"],
        x = ["2025년 총 과업시간", f"T/O 감축 (-{reduced_people}명)", "조직통폐합 및<br>데이터 웹앱 자동화", "2026년 총 과업시간"],
        textposition = "outside",
        text = [
            f"{sum_2025:,.0f}h", 
            f"-{loss_hours:,.0f}h", 
            f"-{efficiency_hours:,.0f}h", 
            f"{sum_2026:,.0f}h"
        ],
        y = [sum_2025, -loss_hours, -efficiency_hours, sum_2026],
        connector = {"line":{"color":"rgb(63, 63, 63)", "width": 2}},
        decreasing = {"marker":{"color":"#FF5A5F"}},
        totals = {"marker":{"color":"#00A699"}}
    ))

    fig_waterfall.update_layout(height=450, showlegend=False)
    st.plotly_chart(fig_waterfall, use_container_width=True)

    # =====================================================================
    # 시각화 2: 2025 vs 2026 직무별 비중 비교 (도넛 차트)
    # =====================================================================
    st.subheader("📊 2. 연도별 직무 투입 시간 비교")
    col1, col2 = st.columns(2)
    
    # 2025년 직무 그룹
    df_job_25 = df_2025.groupby('직무')['과업수행시간'].sum().reset_index()
    # 2026년 직무 그룹
    df_job_26 = df_2026.groupby('직무')['과업수행시간'].sum().reset_index()
    
    with col1:
        st.markdown("**[ 2025년 직무 비중 ]**")
        fig_pie25 = px.pie(df_job_25, values='과업수행시간', names='직무', hole=0.4)
        fig_pie25.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_pie25, use_container_width=True)
        
    with col2:
        st.markdown("**[ 2026년 직무 비중 ]**")
        fig_pie26 = px.pie(df_job_26, values='과업수행시간', names='직무', hole=0.4)
        fig_pie26.update_traces(textposition='inside', textinfo='percent+label')
        st.plotly_chart(fig_pie26, use_container_width=True)

    # =====================================================================
    # 데이터 테이블 비교
    # =====================================================================
    st.subheader("📋 3. 세부 직무 데이터 조회")
    tab1, tab2 = st.tabs(["2026년 현재", "2025년 작년"])
    
    with tab1:
        st.dataframe(df_2026[['직무', '책무', '과업', '과업 담당자', '과업수행시간']], use_container_width=True)
    with tab2:
        st.dataframe(df_2025[['직무', '책무', '과업', '과업 담당자', '과업수행시간']], use_container_width=True)

else:
    st.error("지정된 엑셀 파일을 찾을 수 없습니다. 파일명이 코드와 동일한지, 같은 폴더에 있는지 확인해주세요.")
