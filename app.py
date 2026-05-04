import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

# 페이지 기본 설정
st.set_page_config(page_title="마케팅팀 직무분석", layout="wide")

st.title("📊 마케팅팀 직무분석")
st.markdown("""
**작성자:** 최한엽 (Han)  
본 대시보드는 2025년 대비 2026년 마케팅팀의 인원 변동(-2명) 및 조직 개편에 따른 **업무 효율화 성과**를 시각화한 자료입니다.
""")

# 1. 파일 업로드 기능
uploaded_file = st.sidebar.file_uploader("직무현황표 엑셀 파일을 업로드하세요", type=["xlsx"])

if uploaded_file is not None:
    # 2. 데이터 로드 및 전처리
    @st.cache_data
    def load_data(file):
        # '직무표' 시트 불러오기 (헤더는 2번째 줄에 위치)
        df = pd.read_excel(file, sheet_name='직무표', header=1) 
        
        # 컬럼명 단순화
        df.rename(columns={
            '직무\n(Job)': '직무',
            '책무\n(Duty)': '책무',
            '과업\n(Task)': '과업',
            '과업\n수행시간\n(연간)': '수행시간'
        }, inplace=True)
        
        # 직무, 책무 컬럼의 병합된 빈칸 채우기(Forward Fill)
        df['직무'] = df['직무'].ffill()
        df['책무'] = df['책무'].ffill()
        
        # 수행시간이 숫자인 행만 필터링
        df['수행시간'] = pd.to_numeric(df['수행시간'], errors='coerce')
        df = df.dropna(subset=['수행시간'])
        
        return df

    try:
        df = load_data(uploaded_file)
        
        # 총 수행시간 계산
        total_hours_2026 = df['수행시간'].sum()
        # 1인당 표준근무가능시간
        std_hours_per_person = 1826.7
        # 2명 감소분
        reduced_hours = std_hours_per_person * 2
        
        st.success("데이터 로드 완료!")

        # =====================================================================
        # 시각화 1: 워터폴 차트 (조직개편 및 인원 감축 대응 성과)
        # =====================================================================
        st.subheader("1. 업무 가용시간 변동 및 효율화 성과")
        
        st.markdown("""
        인원 감축(-2명)으로 인한 가용시간 감소분을 **중복업무 통폐합** 및 **데이터 분석 기반 웹앱 자동화**를 통해 상쇄하고, 신규 핵심 사업(연료전지 등)에 여력을 집중하고 있습니다.
        """)

        # 워터폴 차트 데이터 구성
        fig_waterfall = go.Figure(go.Waterfall(
            name = "시간 변동",
            orientation = "v",
            measure = ["relative", "relative", "relative", "relative", "total"],
            x = ["25년 가용시간<br>(인원감축 전)", "T/O 감축<br>(-2명)", "업무 통폐합 및<br>자동화 효율화", "고부가가치<br>업무 재투자", "26년 현재<br>총 가용시간"],
            textposition = "outside",
            text = [
                f"{total_hours_2026 + reduced_hours:,.0f}h", 
                f"-{reduced_hours:,.0f}h", 
                f"+{reduced_hours:,.0f}h", 
                "효율화 시간 전환", 
                f"{total_hours_2026:,.0f}h"
            ],
            y = [total_hours_2026 + reduced_hours, -reduced_hours, reduced_hours, 0, total_hours_2026],
            connector = {"line":{"color":"rgb(63, 63, 63)"}},
            decreasing = {"marker":{"color":"#FF5A5F"}},
            increasing = {"marker":{"color":"#00A699"}},
            totals = {"marker":{"color":"#484848"}}
        ))

        fig_waterfall.update_layout(
            title="연간 마케팅팀 업무시간 변동 추이 (1인 1,826.7h 기준)",
            showlegend=False,
            height=500
        )
        st.plotly_chart(fig_waterfall, use_container_width=True)

        # =====================================================================
        # 시각화 2: 2026년 직무별 투입 시간 비중 (도넛 차트 & 트리맵)
        # =====================================================================
        st.subheader("2. 2026년 마케팅팀 직무별 투입 비중")
        
        col1, col2 = st.columns(2)
        
        # 직무별 그룹화
        df_job = df.groupby('직무')['수행시간'].sum().reset_index()
        
        with col1:
            st.markdown("**직무별 투입시간 (도넛 차트)**")
            fig_pie = px.pie(df_job, values='수행시간', names='직무', hole=0.4,
                             color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_pie.update_traces(textposition='inside', textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)
            
        with col2:
            st.markdown("**세부 과업별 투입시간 (트리맵)**")
            # 트리맵을 위해 0인 데이터 제외
            df_tree = df[df['수행시간'] > 0]
            fig_tree = px.treemap(df_tree, path=['직무', '책무'], values='수행시간',
                                  color='수행시간', color_continuous_scale='Blues')
            st.plotly_chart(fig_tree, use_container_width=True)

        # =====================================================================
        # 데이터 테이블 확인
        # =====================================================================
        st.subheader("3. 세부 직무 데이터 확인")
        job_list = df['직무'].dropna().unique()
        selected_jobs = st.multiselect("확인하고 싶은 직무를 선택하세요:", options=job_list, default=job_list)
        
        df_filtered = df[df['직무'].isin(selected_jobs)]
        st.dataframe(df_filtered[['직무', '책무', '과업', '수행시간', '과업 담당자']], use_container_width=True)

    except Exception as e:
        st.error(f"데이터를 처리하는 중 오류가 발생했습니다. 엑셀 파일 형식을 다시 확인해 주세요.\n오류내용: {e}")

else:
    st.info("👈 왼쪽 사이드바에서 '직무현황표' 엑셀 파일을 업로드해 주세요.")
