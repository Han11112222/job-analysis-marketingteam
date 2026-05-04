import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import os

# 1. 페이지 설정
st.set_page_config(page_title="마케팅팀 직무분석 (ERRC)", layout="wide")

st.title("📊 마케팅팀 직무 및 적정 인원 분석 (2025 vs 2026)")
st.markdown("""
**작성자:** 최한엽 (Han)  
본 대시보드는 2025년 대비 2026년 마케팅팀의 **실제 과업 수행 시간**을 바탕으로 **적정 필요 인원**을 산출하고,  
조직 개편(마케팅부팀 통폐합)에 따른 직무 변화를 **ERRC(제거/감소/증가/창조)** 관점에서 분석한 자료입니다.
""")

# 2. 데이터 로딩 함수
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path, sheet_name='직무표')
    
    df = df.rename(columns={
        '직무\n(Job)': '직무',
        '책무\n(Duty)': '책무',
        '과업\n(Task)': '과업',
        '과업\n수행시간\n(연간)': '과업수행시간'
    })
    
    df = df.loc[:, ~df.columns.duplicated()]
    df['직무'] = df['직무'].ffill()
    df['책무'] = df['책무'].ffill()
    df['과업수행시간'] = pd.to_numeric(df['과업수행시간'], errors='coerce')
    return df.dropna(subset=['과업수행시간'])

# 3. 데이터 불러오기 및 계산
file_2025 = "직무현황표_20250515_2025B_마케팅팀.xlsx"
file_2026 = "직무현황표_20260408_2026A_마케팅팀.xlsx"

if os.path.exists(file_2025) and os.path.exists(file_2026):
    df_2025 = load_data(file_2025)
    df_2026 = load_data(file_2026)
    
    # 총 수행시간 계산
    sum_2025 = df_2025['과업수행시간'].sum()
    sum_2026 = df_2026['과업수행시간'].sum()
    
    # 1인당 표준근무가능시간 기준 필요 인원 산출
    std_hours = 1826.7
    req_ppl_2025 = sum_2025 / std_hours
    req_ppl_2026 = sum_2026 / std_hours
    
    st.divider()

    # =====================================================================
    # 분석 1: 적정 필요 인원 도출 (논리 전개)
    # =====================================================================
    st.subheader("💡 1. 총 과업시간 기반 적정 필요 인원 분석")
    
    st.error(f"""
    **[핵심 요약] 현재 마케팅팀은 인력 운용이 매우 타이트한 한계 상황입니다.**
    * 2026년 기준 마케팅팀의 연간 총 과업시간은 **{sum_2026:,.0f}시간**입니다.
    * 1인당 표준근무가능시간({std_hours:,.1f}시간)으로 환산 시, 현재 산적한 업무를 소화하기 위한 **적정 필요 인원은 {req_ppl_2026:.1f}명**입니다. 
    * 단순 인원 감축의 충격을 방어하기 위해 조직을 통폐합하고 일부 업무를 효율화했음에도 불구하고, 여전히 높은 노동 강도가 요구되고 있습니다.
    """)

    col1, col2 = st.columns(2)
    with col1:
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=['2025년', '2026년'], y=[sum_2025, sum_2026], 
                                 text=[f"{sum_2025:,.0f}h", f"{sum_2026:,.0f}h"], textposition='auto', marker_color=['#1f77b4', '#00A699']))
        fig_bar.update_layout(title="연간 총 과업 수행시간 비교", height=350)
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col2:
        fig_line = go.Figure()
        fig_line.add_trace(go.Bar(x=['2025년 기준', '2026년 기준'], y=[req_ppl_2025, req_ppl_2026],
                                   text=[f"필요 {req_ppl_2025:.1f}명", f"필요 {req_ppl_2026:.1f}명"], textposition='auto', marker_color=['#ff7f0e', '#FF5A5F']))
        fig_line.update_layout(title="표준근무시간 기준 적정 필요 인원", height=350)
        st.plotly_chart(fig_line, use_container_width=True)

    st.divider()

    # =====================================================================
    # 분석 2 & 3: ERRC 관점의 직무 변화 (책무별 증감 비교)
    # =====================================================================
    st.subheader("🎯 2 & 3. ERRC 관점의 직무 변화 및 시사점")
    
    st.markdown("""
    업무 과부하 속에서도 마케팅 1, 2팀을 통합하며 중복 업무를 **제거(E)/감소(R)**하고, 
    미래 성장 동력과 업무 자동화를 위한 역량을 **증가(R)/창조(C)**하는 체질 개선을 병행하고 있습니다.
    """)

    # 책무별 시간 비교 데이터프레임 생성
    duty_2025 = df_2025.groupby('책무')['과업수행시간'].sum()
    duty_2026 = df_2026.groupby('책무')['과업수행시간'].sum()
    
    diff_df = pd.DataFrame({'2025년(h)': duty_2025, '2026년(h)': duty_2026}).fillna(0)
    diff_df['증감(h)'] = diff_df['2026년(h)'] - diff_df['2025년(h)']
    diff_df = diff_df.sort_values('증감(h)')

    # ERRC 분류 로직 (가상의 기준 적용, 데이터에 맞게 수정 가능)
    # - 제거(Eliminate): 2025년엔 있었으나 2026년에 0이 되거나 대폭 축소된 중복 업무
    # - 감소(Reduce): 효율화를 통해 시간을 줄인 업무
    # - 증가(Raise): 전략적으로 투자를 늘린 업무
    # - 창조(Create): 새롭게 생성된 과업 (데이터 분석, 자동화 웹앱 등)

    col_e, col_r, col_ra, col_c = st.columns(4)
    
    with col_e:
        st.error("🗑️ **Eliminate (제거)**\n\n부팀 통합에 따른 중복/파편화 업무 제거")
        # 예: 업무용 마케팅, 영업용 마케팅 등 통합 전 책무명
        st.write("- 업무용 마케팅 (-2,432h)")
        st.write("- 영업용 마케팅 (-1,988h)")
        
    with col_r:
        st.warning("📉 **Reduce (감소)**\n\n효율화 및 선택과 집중을 통한 시간 단축")
        st.write("- 공동주택공급관리 (-1,594h)")
        st.write("- 연료전지마케팅 (-1,742h)")
        
    with col_ra:
        st.success("📈 **Raise (증가)**\n\n마케팅 통합 및 핵심 전략에 역량 집중")
        # 예: 통합된 '업무(영업)용 마케팅' 책무
        st.write("+ 업무(영업)용 마케팅 (+3,194h)")
        st.write("+ 마케팅전략기획 (+230h)")
        
    with col_c:
        st.info("💡 **Create (창조)**\n\n업무 효율 극대화를 위한 신규 시스템 구축")
        st.write("+ **데이터 분석 웹앱 개발 및 배포 (+540h)**")
        st.write("+ **시각화 결과 공유/유지보수**")

    st.markdown("---")
    
    # 데이터 표 (증감 순 정렬)
    st.markdown("**[ 책무별 연간 투입 시간 증감 현황 ]**")
    
    # 색상 적용 함수
    def color_negative_red(val):
        color = 'red' if val < 0 else 'blue' if val > 0 else 'black'
        return f'color: {color}'
        
    st.dataframe(diff_df.style.applymap(color_negative_red, subset=['증감(h)']).format("{:,.0f}"), use_container_width=True)

    # 향후 집중 방향 결론
    st.info("""
    🚀 **[Conclusion & Action Plan]**
    현재 한정된 인력(타이트한 T/O)으로 방대한 과업을 수행하기 위해, 마케팅팀은 **수작업(엑셀 등)에 의존하던 분석 및 관리 업무를 파이썬 웹앱 기반으로 완전히 전환(Create)**하고 있습니다. 
    이를 통해 확보된 시간은 대성에너지의 핵심 이익과 직결되는 **'마케팅전략기획'**과 **'대규모 영업(업무/영업용 통합)'** 분야에 집중적으로 재투자(Raise)할 계획입니다.
    """)

else:
    st.warning("지정된 엑셀 파일을 찾을 수 없습니다. 파일명이 깃허브 코드와 동일한지 확인해주세요.")
