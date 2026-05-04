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
    
    # --- 팀리더 시간 분리 계산 ---
    leader_2025 = df_2025[df_2025['직무'] == '팀리더']['과업수행시간'].sum()
    leader_2026 = df_2026[df_2026['직무'] == '팀리더']['과업수행시간'].sum()
    
    # 팀원(실무진) 순수 과업 시간 계산
    staff_2025_sum = df_2025[df_2025['직무'] != '팀리더']['과업수행시간'].sum()
    staff_2026_sum = df_2026[df_2026['직무'] != '팀리더']['과업수행시간'].sum()
    
    # 1인당 표준근무가능시간 기준 필요 인원 산출 (팀원 기준)
    std_hours = 1826.7
    req_ppl_2025 = staff_2025_sum / std_hours
    req_ppl_2026 = staff_2026_sum / std_hours
    
    st.divider()

    # =====================================================================
    # 분석 1: 적정 필요 인원 도출 (팀리더 분리)
    # =====================================================================
    st.subheader("💡 1. 실무진(팀원) 과업시간 기반 적정 필요 인원 분석")
    
    st.error(f"""
    **[핵심 요약] 마케팅팀 실무 인력 운용 한계 도달 및 타이트한 T/O 증명**
    * **고정 인력:** 팀리더 1명 (연간 {leader_2026:,.0f}h 수행)은 고정 산정합니다.
    * 2026년 기준 실무진(팀원)이 감당해야 할 순수 총 과업시간은 **{staff_2026_sum:,.0f}시간**입니다.
    * 1인당 표준근무가능시간({std_hours:,.1f}시간)으로 환산 시, 해당 업무를 소화하기 위한 **실무진 적정 필요 인원은 {req_ppl_2026:.1f}명**입니다. 
    * 현재 팀원 배정(10명)은 필요 인원 대비 **{req_ppl_2026 - 10:.1f}명 부족**한 상태이며, 웹앱 자동화 같은 극한의 효율화 없이는 정상적인 업무 수행이 불가능한 수치입니다.
    """)

    col1, col2 = st.columns(2)
    with col1:
        # 총 과업시간(합산) 계산
        total_2025 = staff_2025_sum + leader_2025
        total_2026 = staff_2026_sum + leader_2026
        
        fig_bar = go.Figure()
        # textfont 옵션을 추가하여 폰트 크기 통일 (예: 14)
        fig_bar.add_trace(go.Bar(name='팀원(실무)', x=['2025년', '2026년'], y=[staff_2025_sum, staff_2026_sum], text=[f"{staff_2025_sum:,.0f}h", f"{staff_2026_sum:,.0f}h"], textposition='inside', marker_color='#00A699', textfont=dict(size=14, color='white')))
        # 팀리더 막대의 폰트 크기도 14로 고정
        fig_bar.add_trace(go.Bar(name='팀리더', x=['2025년', '2026년'], y=[leader_2025, leader_2026], text=[f"{leader_2025:,.0f}h", f"{leader_2026:,.0f}h"], textposition='inside', marker_color='#1f77b4', textfont=dict(size=14, color='white')))
        
        # 막대 최상단에 전체 합산 시간을 [ *****h ] 형태로 추가
        fig_bar.add_trace(go.Scatter(
            x=['2025년', '2026년'], 
            y=[total_2025, total_2026],
            text=[f"<b>[ {total_2025:,.0f}h ]</b>", f"<b>[ {total_2026:,.0f}h ]</b>"],
            mode='text',
            textposition='top center',
            showlegend=False,
            hoverinfo='skip'
        ))
        
        # 글자가 상단에 잘리지 않도록 margin(여백) t=60 옵션 추가
        fig_bar.update_layout(title="연간 과업 수행시간 비교 (누적)", barmode='stack', height=400, margin=dict(t=60))
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col2:
        fig_line = go.Figure()
        fig_line.add_trace(go.Bar(x=['2025년 기준', '2026년 기준'], y=[req_ppl_2025, req_ppl_2026],
                                   text=[f"팀원 필요 {req_ppl_2025:.1f}명", f"팀원 필요 {req_ppl_2026:.1f}명"], textposition='auto', marker_color=['#ff7f0e', '#FF5A5F']))
        fig_line.add_hline(y=10, line_dash="dash", line_color="red", annotation_text="현재 팀원 정원 (10명)", annotation_position="bottom right")
        # 제목 (팀원 한정) -> (팀장 제외)로 수정
        fig_line.update_layout(title="표준근무시간 기준 적정 필요 인원 (팀장 제외)", height=400)
        st.plotly_chart(fig_line, use_container_width=True)

    st.divider()

    # =====================================================================
    # 분석 2 & 3: ERRC 관점의 직무 변화 (책무별 증감 비교)
    # =====================================================================
    st.subheader("🎯 2 & 3. ERRC 관점의 직무 변화 및 세부 과업 비교")
    
    st.markdown("""
    업무 과부하 속에서도 마케팅 1, 2팀을 통합하며 중복 업무를 **제거(E)/감소(R)**하고, 
    미래 성장 동력과 업무 자동화를 위한 역량을 **증가(R)/창조(C)**하는 체질 개선을 병행하고 있습니다.
    """)

    # 책무별 시간 비교 데이터프레임 생성
    duty_2025 = df_2025.groupby('책무')['과업수행시간'].sum()
    duty_2026 = df_2026.groupby('책무')['과업수행시간'].sum()
    diff_df = pd.DataFrame({'2025년(h)': duty_2025, '2026년(h)': duty_2026}).fillna(0)
    diff_df['증감(h)'] = diff_df['2026년(h)'] - diff_df['2025년(h)']
    
    # ERRC 카드 섹션
    col_e, col_r, col_ra, col_c = st.columns(4)
    
    with col_e:
        st.error("🗑️ **Eliminate (제거)**\n\n부팀 통합에 따른 중복 업무 제거")
        st.write("- 업무용 마케팅 (-2,432h)")
        st.write("- 영업용 마케팅 (-1,988h)")
        
    with col_r:
        st.warning("📉 **Reduce (감소)**\n\n효율화 및 선택과 집중을 통한 시간 단축")
        st.write("- 공동주택공급관리 (-1,594h)")
        st.write("- 연료전지마케팅 (-1,742h)")
        
    with col_ra:
        st.success("📈 **Raise (증가)**\n\n마케팅 통합 및 핵심 전략에 역량 집중")
        st.write("+ 업무(영업)용 마케팅 (+3,194h)")
        st.write("+ 마케팅전략기획 (+230h)")
        
    with col_c:
        st.info("💡 **Create (창조)**\n\n업무 효율 극대화를 위한 신규 시스템 구축")
        st.write("+ **데이터 분석 웹앱 개발/배포 (+540h)**")
        st.write("+ **시각화 결과 공유/유지보수**")

    # ERRC 하단 직관적인 막대그래프
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**[ 주요 책무별 증감 현황 그래프 ]**")
    
    # 증감이 0이 아닌 것만 필터링 후 정렬
    diff_plot_df = diff_df[diff_df['증감(h)'] != 0].sort_values('증감(h)')
    
    fig_errc_bar = px.bar(diff_plot_df, x='증감(h)', y=diff_plot_df.index, orientation='h',
                          color='증감(h)', color_continuous_scale=px.colors.diverging.RdBu)
    fig_errc_bar.update_layout(height=500, showlegend=False)
    st.plotly_chart(fig_errc_bar, use_container_width=True)

    st.markdown("---")
    
    # 과업별 세부 증감 내용 (아코디언 형태)
    st.markdown("**[ 세부 과업별 가감 내역 확인 ]**")
    
    # 과업 레벨 병합
    task_25 = df_2025.groupby(['책무', '과업'])['과업수행시간'].sum().reset_index()
    task_26 = df_2026.groupby(['책무', '과업'])['과업수행시간'].sum().reset_index()
    merged_task = pd.merge(task_25, task_26, on=['책무', '과업'], how='outer', suffixes=('_2025', '_2026')).fillna(0)
    merged_task['증감(h)'] = merged_task['과업수행시간_2026'] - merged_task['과업수행시간_2025']
    merged_task = merged_task[merged_task['증감(h)'] != 0]

    with st.expander("🔻 가장 많이 감소/제거된 과업 확인하기"):
        st.dataframe(merged_task.sort_values('증감(h)').head(10)[['책무', '과업', '증감(h)', '과업수행시간_2025', '과업수행시간_2026']].style.format("{:,.0f}", subset=['증감(h)', '과업수행시간_2025', '과업수행시간_2026']), use_container_width=True)

    with st.expander("🔺 가장 많이 증가/창조된 과업 확인하기"):
        st.dataframe(merged_task.sort_values('증감(h)', ascending=False).head(10)[['책무', '과업', '증감(h)', '과업수행시간_2025', '과업수행시간_2026']].style.format("{:,.0f}", subset=['증감(h)', '과업수행시간_2025', '과업수행시간_2026']), use_container_width=True)

    # 향후 집중 방향 결론
    st.info(f"""
    🚀 **[Conclusion & Action Plan]**
    현재 마케팅팀은 **필요 인원({req_ppl_2026:.1f}명) 대비 실제 인원(10명)** 이라는 극도로 타이트한 환경 속에서 일하고 있습니다. 
    이를 극복하기 위해 기존에 수작업으로 진행되던 반복 분석 및 관리 업무를 **파이썬 기반 데이터 분석 웹앱으로 전환(Create)**하여 물리적 한계를 시스템으로 돌파하고 있습니다.
    
    이렇게 자동화를 통해 확보된 여력은 대성에너지의 미래 먹거리인 **'연료전지(HPS) 등 신규 사업'**과 핵심 이익 분야인 **'통합 대규모 영업(업무/영업용)'** 및 **'마케팅전략기획'**에 집중적으로 투입(Raise)하여, 적은 인원으로도 최대의 비즈니스 임팩트를 창출해 나갈 것입니다.
    """)

else:
    st.warning("지정된 엑셀 파일을 찾을 수 없습니다. 파일명이 깃허브 코드와 동일한지 확인해주세요.")
