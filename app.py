import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import os

# 1. 페이지 설정
st.set_page_config(page_title="마케팅팀 직무분석 (ERRC)", layout="wide")

st.title("📊 마케팅팀 직무 및 적정 인원 분석 (2025 vs 2026)")
st.markdown("""
본 대시보드는 2025년 대비 2026년 마케팅팀의 **실제 과업 수행 시간**을 바탕으로 **적정 필요 인원**을 산출하고,  
조직 개편(마케팅부팀 통폐합)에 따른 직무 변화를 **ERRC(제거/감소/증가/창조)** 관점에서 분석한 자료입니다.
""")

# 2. 데이터 로딩 및 전처리 함수 (헤더 2줄 처리)
@st.cache_data
def load_data(file_path):
    # 멀티 인덱스(헤더가 2줄)로 읽어오기
    df = pd.read_excel(file_path, sheet_name='직무표', header=[0, 1])
    
    # 두 줄의 헤더를 하나로 병합 (Unnamed 제거)
    df.columns = ['_'.join([str(c) for c in col if "Unnamed" not in str(c)]).strip('_').replace('\n', '') for col in df.columns]
    
    # 컬럼명 단순화 매핑
    df = df.rename(columns={
        '직무(Job)': '직무명',
        '책무(Duty)': '책무명',
        '과업(Task)': '과업명',
        '과업수행시간(연간)': '수행시간',
        '업무 수준 및 등급_중요도(1~5)': '중요도',
        '업무 수준 및 등급_난이도(1~5)': '난이도',
        '업무 수준 및 등급_숙련도(1~5)': '숙련도',
        '업무 수준 및 등급_등급': '업무등급'
    })
    
    df = df.loc[:, ~df.columns.duplicated()]
    df['직무명'] = df['직무명'].ffill()
    df['책무명'] = df['책무명'].ffill()
    df['수행시간'] = pd.to_numeric(df['수행시간'], errors='coerce')
    
    # 중요도, 난이도, 숙련도 숫자형 변환
    for col in ['중요도', '난이도', '숙련도']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
            
    return df.dropna(subset=['수행시간'])

# 3. 데이터 불러오기 및 계산
file_2025 = "직무현황표_20250515_2025B_마케팅팀.xlsx"
file_2026 = "직무현황표_20260408_2026A_마케팅팀.xlsx"

if os.path.exists(file_2025) and os.path.exists(file_2026):
    df_2025 = load_data(file_2025)
    df_2026 = load_data(file_2026)
    
    # --- 팀리더 시간 분리 계산 ---
    leader_2025 = df_2025[df_2025['직무명'] == '팀리더']['수행시간'].sum()
    leader_2026 = df_2026[df_2026['직무명'] == '팀리더']['수행시간'].sum()
    
    # 팀원(실무진) 순수 과업 시간 계산
    staff_2025_sum = df_2025[df_2025['직무명'] != '팀리더']['수행시간'].sum()
    staff_2026_sum = df_2026[df_2026['직무명'] != '팀리더']['수행시간'].sum()
    
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
        total_2025 = staff_2025_sum + leader_2025
        total_2026 = staff_2026_sum + leader_2026
        
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(name='팀원(실무)', x=['2025년', '2026년'], y=[staff_2025_sum, staff_2026_sum], text=[f"{staff_2025_sum:,.0f}h", f"{staff_2026_sum:,.0f}h"], textposition='inside', marker_color='#00A699', textfont=dict(size=14, color='white'), constraintext='none'))
        fig_bar.add_trace(go.Bar(name='팀리더', x=['2025년', '2026년'], y=[leader_2025, leader_2026], text=[f"{leader_2025:,.0f}h", f"{leader_2026:,.0f}h"], textposition='inside', marker_color='#1f77b4', textfont=dict(size=14, color='white'), constraintext='none'))
        
        fig_bar.add_trace(go.Scatter(
            x=['2025년', '2026년'], y=[total_2025, total_2026],
            text=[f"<b>[ {total_2025:,.0f}h ]</b>", f"<b>[ {total_2026:,.0f}h ]</b>"],
            mode='text', textposition='top center', showlegend=False, hoverinfo='skip'
        ))
        fig_bar.update_layout(title="연간 과업 수행시간 비교 (누적)", barmode='stack', height=400, margin=dict(t=60))
        st.plotly_chart(fig_bar, use_container_width=True)
        
    with col2:
        fig_line = go.Figure()
        fig_line.add_trace(go.Bar(x=['2025년 기준', '2026년 기준'], y=[req_ppl_2025, req_ppl_2026],
                                   text=[f"팀원 필요 {req_ppl_2025:.1f}명", f"팀원 필요 {req_ppl_2026:.1f}명"], textposition='auto', marker_color=['#ff7f0e', '#FF5A5F']))
        fig_line.add_hline(y=10, line_dash="dash", line_color="red", annotation_text="현재 팀원 정원 (10명)", annotation_position="bottom right")
        fig_line.update_layout(title="표준근무시간 기준 적정 필요 인원 (팀장 제외)", height=400)
        st.plotly_chart(fig_line, use_container_width=True)

    st.divider()

    # =====================================================================
    # 분석 2: ERRC 관점의 직무 변화 (책무별 증감 비교)
    # =====================================================================
    st.subheader("🎯 2. ERRC 관점의 직무 변화 및 세부 과업 비교")
    
    st.markdown("""
    업무 과부하 속에서도 마케팅 1, 2팀을 통합하며 중복 업무를 **제거(E)/감소(R)**하고, 
    미래 성장 동력과 업무 자동화를 위한 역량을 **증가(R)/창조(C)**하는 체질 개선을 병행하고 있습니다.
    """)

    duty_2025 = df_2025.groupby('책무명')['수행시간'].sum()
    duty_2026 = df_2026.groupby('책무명')['수행시간'].sum()
    diff_df = pd.DataFrame({'2025년(h)': duty_2025, '2026년(h)': duty_2026}).fillna(0)
    diff_df['증감(h)'] = diff_df['2026년(h)'] - diff_df['2025년(h)']
    
    col_e, col_r, col_ra, col_c = st.columns(4)
    
    with col_e:
        st.error("🗑️ **Eliminate (제거)**\n\n부팀 통합에 따른 중복 파편화 업무 제거")
        st.write("- 업무용 마케팅 (-2,432h)")
        st.write("- 영업용 마케팅 (-1,988h)")
        
    with col_r:
        st.warning("📉 **Reduce (감소)**\n\n자동화 및 선택과 집중을 통한 시간 단축")
        st.write("- 연료전지마케팅 (-1,742h)")
        st.write("- 공동주택공급관리 (-1,594h)")
        st.write("- 산업용마케팅 (-666h)")
        
    with col_ra:
        st.success("📈 **Raise (증가)**\n\n핵심 전략 및 기획에 역량 집중")
        st.write("+ 마케팅전략기획 (+230h)")
        st.write("  *(영업활성화 과업 +690h 증가)*")
        
    with col_c:
        st.info("💡 **Create (창조)**\n\n기능 통합형 신규 마케팅 체계 구축")
        st.write("+ 업무(영업)용 마케팅 (+3,194h)")
        st.write("  *(업무/영업용 일원화 신설)*")

    # 향후 집중 방향 결론 
    st.info(f"""
    🚀 **[Conclusion & Action Plan]**
    현재 마케팅팀은 **필요 인원({req_ppl_2026:.1f}명) 대비 실제 인원(10명)** 이라는 극도로 타이트한 환경 속에서 일하고 있습니다. 
    이를 극복하기 위해 기존에 막대한 시간이 소요되던 분석 및 관리 업무(공동주택/연료전지 등)를 **데이터 분석 기반 파이썬 웹앱으로 자동화**하여 획기적으로 감축(Reduce) 시켰습니다.
    
    이렇게 자동화를 통해 시스템적으로 방어한 여력은, 대성에너지의 핵심 이익 분야인 **'통합 대규모 영업(업무/영업용 마케팅 신설)'**과 **'마케팅전략기획(영업활성화)'**에 집중적으로 투입(Raise/Create)하여 적은 인원으로도 최대의 비즈니스 임팩트를 창출해 나갈 것입니다.
    """)

    st.divider()

    # =====================================================================
    # 분석 3: 책무별 업무량 분석 표 및 비교 차트 (신규 추가)
    # =====================================================================
    st.subheader("📊 3. 책무별 업무량 정밀 분석 및 연도별 비교")
    
    st.markdown("현재 2026년 기준 직무/책무별 세부 수준(중요도, 난이도, 숙련도, 등급)과 업무량 구성비를 확인하실 수 있습니다.")
    
    # 2026년 책무별 분석표 생성
    agg_26 = df_2026.groupby(['직무명', '책무명']).agg(
        과업수=('과업명', 'count'),
        연간수행시간=('수행시간', 'sum'),
        중요도=('중요도', 'mean'),
        난이도=('난이도', 'mean'),
        숙련도=('숙련도', 'mean')
    ).reset_index()

    total_hours_26 = agg_26['연간수행시간'].sum()
    agg_26['업무량 구성비(%)'] = (agg_26['연간수행시간'] / total_hours_26 * 100).round(1)

    # 최빈값(가장 많이 등장하는 등급) 구하기
    def get_mode(x):
        return x.mode().iloc[0] if not x.mode().empty else None

    grade_26 = df_2026.groupby(['직무명', '책무명'])['업무등급'].agg(get_mode).reset_index()
    agg_26 = pd.merge(agg_26, grade_26, on=['직무명', '책무명'], how='left')

    # 소수점 첫째자리 반올림
    agg_26['중요도'] = agg_26['중요도'].round(1)
    agg_26['난이도'] = agg_26['난이도'].round(1)
    agg_26['숙련도'] = agg_26['숙련도'].round(1)

    # 컬럼 순서 정렬 및 출력
    cols = ['직무명', '책무명', '과업수', '연간수행시간', '업무량 구성비(%)', '중요도', '난이도', '숙련도', '업무등급']
    agg_26 = agg_26[cols].sort_values(['직무명', '연간수행시간'], ascending=[True, False])
    
    # 데이터프레임 시각화 (인덱스 숨김)
    st.dataframe(agg_26.style.format({
        '연간수행시간': '{:,.1f}',
        '중요도': '{:.1f}', '난이도': '{:.1f}', '숙련도': '{:.1f}'
    }), hide_index=True, use_container_width=True)

    st.markdown("<br>**[ 2025년 vs 2026년 책무별 수행시간 변동 비교 ]**", unsafe_allow_html=True)
    
    # 두 연도 비교를 위한 데이터프레임 병합
    comp_df = diff_df.reset_index()
    
    # 증감이 0이 아닌 유의미한 항목만 필터링하여 시각화
    comp_df = comp_df[(comp_df['2025년(h)'] > 0) | (comp_df['2026년(h)'] > 0)]
    
    # Plotly Grouped Bar Chart 작성
    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(
        x=comp_df['2025년(h)'], y=comp_df['책무명'],
        name='2025년', orientation='h', marker_color='#1f77b4'
    ))
    fig_comp.add_trace(go.Bar(
        x=comp_df['2026년(h)'], y=comp_df['책무명'],
        name='2026년', orientation='h', marker_color='#00A699'
    ))

    fig_comp.update_layout(
        barmode='group',
        height=600,
        margin=dict(l=150), # 책무명 라벨이 길 경우 잘리지 않도록 좌측 여백 확보
        xaxis_title="연간 과업 수행시간 (Hours)",
        yaxis={'categoryorder':'total ascending'} # 전체 합계가 높은 순서대로 정렬
    )
    st.plotly_chart(fig_comp, use_container_width=True)

else:
    st.warning("지정된 엑셀 파일을 찾을 수 없습니다. 파일명이 깃허브 코드와 동일한지 확인해주세요.")
