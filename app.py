import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import os
import unicodedata

# 1. 페이지 설정
st.set_page_config(page_title="마케팅팀 직무분석 (ERRC)", layout="wide")

st.title("📊 마케팅팀 직무 및 적정 인원 분석 (2026년 상반기 vs 2026년 하반기)")
st.markdown(
    "2026년 상반기 대비 하반기 마케팅팀의 **실제 과업 수행 시간**을 바탕으로 **적정 필요 인원**을 산출합니다.  \n"
    "조직 개편에 따른 직무 변화는 **ERRC(제거/감소/증가/창조)** 관점에서 분석했습니다.  \n"
    "※ 2025년 자료는 참고용으로 함께 표기합니다."
)

st.info("""
**[주요 변화]**
* **기획부팀 :** CS지원팀 업무 일부 추가
* **마케팅부팀 :** (기존) 용도별 담당자 통합 → (현재) 용도별 담당자 분리
""")

# 2. 데이터 로딩 및 전처리 함수 (헤더 2줄 처리)
@st.cache_data
def load_data(file_path):
    df = pd.read_excel(file_path, sheet_name='직무표', header=[0, 1])

    df.columns = ['_'.join([str(c) for c in col if "Unnamed" not in str(c)]).strip('_').replace('\n', '') for col in df.columns]

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
    df['과업명'] = df['과업명'].ffill()
    df['수행시간'] = pd.to_numeric(df['수행시간'], errors='coerce')

    for col in ['중요도', '난이도', '숙련도']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    return df.dropna(subset=['수행시간'])

# 3. 데이터 불러오기 및 계산
# ※ 한글 파일명은 올린 환경에 따라 내부 글자 구성(NFC/NFD)이 달라 정확히 일치하지 않을 수 있어,
#    파일명에 포함된 구분 키워드(_2025B_ / _2026A_ / _2026B_)로 자동 탐색합니다.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def find_file(keyword):
    matches = []
    for f in os.listdir(BASE_DIR):
        n = unicodedata.normalize('NFC', f)
        if n.endswith('.xlsx') and not n.startswith('~$') and keyword in n:
            matches.append(f)
    # 같은 키워드 파일이 여러 개면 파일명이 가장 늦은(최신 날짜) 파일 사용
    return os.path.join(BASE_DIR, sorted(matches)[-1]) if matches else None

file_2025 = find_file('_2025B_')     # 2025년
file_2026A = find_file('_2026A_')    # 2026년 상반기
file_2026B = find_file('_2026B_')    # 2026년 하반기

LBL_25, LBL_A, LBL_B = '2025년', '2026년 상반기', '2026년 하반기'
COLOR_25, COLOR_A, COLOR_B = '#1f77b4', '#00A699', '#8E44AD'

custom_sort_order = ['팀리더', '일반관리', '영업기획/관리', '주택용 수요개발', '주택용 외 수요개발']

# 1번 섹션 그래프 공통 스타일
FONT = dict(family="Pretendard, 'Noto Sans KR', 'Malgun Gothic', sans-serif", size=13, color="#334155")
ACCENT_RED = "#E11D48"
JOB_COLORS = {
    '팀리더': '#334155',
    '일반관리': '#94A3B8',
    '영업기획/관리': '#F59E0B',
    '주택용 수요개발': '#00A699',
    '주택용 외 수요개발': '#6366F1',
}
EXTRA_COLORS = ['#EC4899', '#84CC16', '#06B6D4', '#A855F7']
CHART_H = 600          # 기존 400 → 600 (1.5배)
PLOT_CFG = {'displayModeBar': False}


def chart_title(text, toggle_label=None):
    """그래프 제목 행 (오른쪽 끝에 선택적으로 토글 배치) - 두 그래프의 제목 행 높이를 동일하게 맞춤"""
    t_left, t_right = st.columns([3, 1])
    t_left.markdown(
        f"<div style='min-height:44px;display:flex;align-items:center;font-size:1.1rem;font-weight:700;color:#0F172A;'>{text}</div>",
        unsafe_allow_html=True
    )
    if toggle_label:
        with t_right:
            return st.toggle(toggle_label, value=False, help="켜면 막대가 직무별로 나뉘어 표시됩니다.")
    return False


def style_fig(fig, height=CHART_H):
    """그래프 공통 스타일 (여백/폰트/그리드/범례)"""
    fig.update_layout(
        template="plotly_white", font=FONT, height=height,
        margin=dict(t=90, b=50, l=60, r=130),
        legend=dict(orientation="h", x=0, y=1.0, yanchor="bottom", font=dict(size=12)),
        plot_bgcolor="white", paper_bgcolor="white",
        hoverlabel=dict(font=dict(family=FONT["family"], size=13)),
    )
    fig.update_xaxes(showgrid=False, linecolor="#CBD5E1", ticks="")
    fig.update_yaxes(gridcolor="#EEF2F6", zeroline=False, ticks="", tickfont=dict(color="#64748B"))
    return fig


def add_ref_line(fig, y, label=None):
    """가로 점선 기준선 (+ 필요 시 우측 라벨)"""
    fig.add_hline(y=y, line_dash="dash", line_color=ACCENT_RED, line_width=2)
    if label:
        fig.add_annotation(xref="paper", yref="y", x=1.0, y=y, xanchor="left", xshift=8,
                           text=label, showarrow=False, align="left",
                           font=dict(color=ACCENT_RED, size=12))

if all([file_2025, file_2026A, file_2026B]):
    df_2025 = load_data(file_2025)
    df_2026A = load_data(file_2026A)
    df_2026B = load_data(file_2026B)

    dfs = [df_2025, df_2026A, df_2026B]
    labels = [LBL_25, LBL_A, LBL_B]

    leader = [d[d['직무명'] == '팀리더']['수행시간'].sum() for d in dfs]
    staff = [d[d['직무명'] != '팀리더']['수행시간'].sum() for d in dfs]
    total = [l + s for l, s in zip(leader, staff)]

    std_hours = 1827                    # 1인당 표준근무가능시간(h)
    CUR_PPL = 10                        # 현재 팀원 정원
    cap_hours = CUR_PPL * std_hours     # 현재 팀원 10명 기준 연간 가용시간
    req_ppl = [s / std_hours for s in staff]

    st.divider()

    # =====================================================================
    # 분석 1: 적정 필요 인원 도출 (팀리더 분리)
    # =====================================================================
    st.subheader("💡 1. 실무진(팀원) 과업시간 기반 적정 필요 인원 분석")

    st.error(f"""
**[핵심 요약] 마케팅팀 실무 인력 운용 한계 도달 및 타이트한 T/O 증명**
* **고정 인력:** 팀리더 1명 (연간 {leader[2]:,.0f}h 수행)은 고정 산정합니다.
* 2026년 하반기 기준 실무진(팀원)이 감당해야 할 순수 총 과업시간은 **{staff[2]:,.0f}시간**입니다. (상반기 {staff[1]:,.0f}시간 대비 {staff[2] - staff[1]:+,.0f}시간)
* 1인당 표준근무가능시간({std_hours:,.0f}시간)으로 환산 시, 해당 업무를 소화하기 위한 **실무진 적정 필요 인원은 {req_ppl[2]:.1f}명**입니다. (상반기 {req_ppl[1]:.1f}명)
* 현재 팀원 배정({CUR_PPL}명)은 필요 인원 대비 **{req_ppl[2] - CUR_PPL:.1f}명 부족**한 상태이며, 웹앱 자동화 같은 극한의 효율화 없이는 정상적인 업무 수행이 불가능한 수치입니다.
""")

    col1, col2 = st.columns(2)
    with col1:
        # 제목 행의 오른쪽 끝에 '직무별 구분' 토글 배치
        show_job = chart_title("연간 과업 수행시간 비교 (누적)", toggle_label="직무별 구분")
        fig_bar = go.Figure()
        job_notes = []          # (하반기 막대 내 세그먼트 중앙 높이, 직무별 증감) -> 오른쪽에 숫자로 표시
        cum_b = 0               # 하반기 막대의 누적 높이

        def add_stack(name, vals, color):
            fig_bar.add_trace(go.Bar(
                name=name, x=labels, y=vals, marker_color=color,
                marker_line=dict(color="white", width=1.5),
                text=[f"{v:,.0f}h" if v >= 500 else "" for v in vals],
                textposition="inside", insidetextanchor="middle",
                textfont=dict(size=12, color="white"),
                hovertemplate="%{x}<br>" + name + ": %{y:,.0f}h<extra></extra>"
            ))

        if show_job:
            all_jobs = pd.concat([d['직무명'] for d in dfs]).dropna().unique().tolist()
            job_order = [j for j in custom_sort_order if j != '팀리더' and j in all_jobs]
            job_order += [j for j in all_jobs if j not in custom_sort_order]
            stack_list = []
            extra_i = 0
            for job in job_order:                      # 팀원 직무를 아래부터 쌓고
                vals = [d[d['직무명'] == job]['수행시간'].sum() for d in dfs]
                if sum(vals) == 0:
                    continue
                color = JOB_COLORS.get(job)
                if color is None:
                    color = EXTRA_COLORS[extra_i % len(EXTRA_COLORS)]
                    extra_i += 1
                stack_list.append((job, vals, color))
            stack_list.append(('팀리더', leader, JOB_COLORS['팀리더']))   # 팀리더는 맨 위
            for name, vals, color in stack_list:
                add_stack(name, vals, color)
                job_notes.append((cum_b + vals[2] / 2, vals[2] - vals[1]))   # 하반기 - 상반기
                cum_b += vals[2]
        else:
            add_stack('팀원(실무)', staff, COLOR_A)
            add_stack('팀리더', leader, JOB_COLORS['팀리더'])

        # 합계 라벨
        fig_bar.add_trace(go.Scatter(
            x=labels, y=total, mode="text", showlegend=False, hoverinfo="skip",
            text=[f"<b>{v:,.0f}h</b>" for v in total], textposition="top center",
            textfont=dict(size=14, color="#0F172A")
        ))

        # 현재 팀원 10명 기준 가용시간 점선 (범례 + 라벨)
        fig_bar.add_trace(go.Scatter(
            x=[None], y=[None], mode="lines",
            line=dict(color=ACCENT_RED, dash="dash", width=2),
            name=f"현재 팀원 {CUR_PPL}명 기준 ({cap_hours:,.0f}h)"
        ))

        style_fig(fig_bar)
        fig_bar.update_layout(barmode="stack", bargap=0.35)
        fig_bar.update_yaxes(range=[0, max(total + [cap_hours]) * 1.12], tickformat=",", title_text="연간 수행시간 (h)")

        if show_job:
            # 점선 라벨은 생략(범례에 표시)하고, 오른쪽 여백에 '상반기 대비 직무별 증감'을 숫자로만 표시
            add_ref_line(fig_bar, cap_hours)
            fig_bar.update_layout(margin=dict(t=90, b=50, l=60, r=110))

            def fmt_diff(v):
                return "0h" if abs(v) < 0.5 else f"{v:+,.0f}h"

            def diff_color(v):
                return "#2563EB" if v > 0 else ("#E11D48" if v < 0 else "#64748B")

            fig_bar.add_annotation(xref="paper", yref="paper", x=1.0, y=1.0, xanchor="left", yanchor="top", xshift=10,
                                   text="<b>상반기 대비</b>", showarrow=False, font=dict(size=12, color="#64748B"))
            for y_mid, diff in job_notes:
                fig_bar.add_annotation(xref="paper", yref="y", x=1.0, y=y_mid, xanchor="left", xshift=10,
                                       text=f"<b>{fmt_diff(diff)}</b>", showarrow=False,
                                       font=dict(size=14, color=diff_color(diff)))
            net = sum(d for _, d in job_notes)
            fig_bar.add_annotation(xref="paper", yref="y", x=1.0, y=total[2], xanchor="left", xshift=10, yshift=16,
                                   text=f"<b>합계 {fmt_diff(net)}</b>", showarrow=False,
                                   font=dict(size=14, color=diff_color(net)))
        else:
            add_ref_line(fig_bar, cap_hours, f"팀원 {CUR_PPL}명 기준<br><b>{cap_hours:,.0f}h</b>")

        st.plotly_chart(fig_bar, use_container_width=True, config=PLOT_CFG)

    with col2:
        chart_title("표준근무시간 기준 적정 필요 인원 (팀장 제외)")
        fig_line = go.Figure(go.Bar(
            x=[f"{l} 기준" for l in labels], y=req_ppl,
            marker_color=[COLOR_25, COLOR_A, COLOR_B],
            text=[f"<b>{v:.1f}명</b>" for v in req_ppl], textposition="outside",
            textfont=dict(size=15, color="#0F172A"),
            hovertemplate="%{x}<br>팀원 필요 %{y:.1f}명<extra></extra>"
        ))
        add_ref_line(fig_line, CUR_PPL, f"현재 팀원 정원<br><b>{CUR_PPL}명</b>")
        style_fig(fig_line)
        fig_line.update_layout(bargap=0.4, showlegend=False)
        fig_line.update_yaxes(range=[0, max(req_ppl + [CUR_PPL]) * 1.18], title_text="필요 인원 (명)")
        try:
            fig_line.update_layout(barcornerradius=8)   # 둥근 막대 (plotly 최신 버전)
        except Exception:
            pass
        st.plotly_chart(fig_line, use_container_width=True, config=PLOT_CFG)

    st.divider()

    # =====================================================================
    # 분석 2: ERRC 관점의 직무 변화 (2026 상반기 → 하반기, 책무별 증감 비교)
    # =====================================================================
    st.subheader("🎯 2. ERRC 관점의 직무 변화 및 세부 과업 비교 (2026년 상반기 → 하반기)")

    st.markdown(
        "2026년 상반기 대비 하반기 직무표를 비교하여, 조직 개편에 따라 업무가 "
        "<b>제거(E)/감소(R)</b>되거나 <b>증가(R)/창조(C)</b>된 내역을 책무 단위로 정리했습니다.",
        unsafe_allow_html=True
    )

    duty_2025 = df_2025.groupby('책무명')['수행시간'].sum()
    duty_2026A = df_2026A.groupby('책무명')['수행시간'].sum()
    duty_2026B = df_2026B.groupby('책무명')['수행시간'].sum()
    diff_df = pd.DataFrame({
        '2025년(h)': duty_2025,
        '2026년 상반기(h)': duty_2026A,
        '2026년 하반기(h)': duty_2026B
    }).fillna(0)
    diff_df['증감(h)'] = diff_df['2026년 하반기(h)'] - diff_df['2026년 상반기(h)']   # 상반기 → 하반기

    # ERRC 자동 분류 (상반기 → 하반기)
    a, b, d = diff_df['2026년 상반기(h)'], diff_df['2026년 하반기(h)'], diff_df['증감(h)']
    elim = diff_df[(a > 0) & (b == 0)].sort_values('증감(h)')
    create = diff_df[(a == 0) & (b > 0)].sort_values('증감(h)', ascending=False)
    reduce_ = diff_df[(a > 0) & (b > 0) & (d < -0.5)].sort_values('증감(h)')
    raise_ = diff_df[(a > 0) & (b > 0) & (d > 0.5)].sort_values('증감(h)', ascending=False)

    def render_items(df_items, empty_msg="해당 없음"):
        if df_items.empty:
            st.markdown(f"• {empty_msg}")
            return
        for name, row in df_items.iterrows():
            st.markdown(f"• {name} ({row['증감(h)']:+,.0f}h)")

    col_e, col_r, col_ra, col_c = st.columns(4)

    with col_e:
        st.error("🗑️ **Eliminate (제거)**\n\n상반기 대비 하반기에 사라진 책무")
        render_items(elim)

    with col_r:
        st.warning("📉 **Reduce (감소)**\n\n자동화 및 선택과 집중을 통한 시간 단축")
        render_items(reduce_)

    with col_ra:
        st.success("📈 **Raise (증가)**\n\n추가 업무 반영 및 핵심 역량 집중")
        render_items(raise_)

    with col_c:
        st.info("💡 **Create (창조)**\n\n하반기에 새로 신설된 책무")
        render_items(create)

    # 과업 단위 변경 내역 (상반기 → 하반기)
    with st.expander("🔍 과업 단위 세부 변경 내역 보기 (상반기 → 하반기)"):
        task_a = df_2026A.groupby(['책무명', '과업명'])['수행시간'].sum().rename('상반기(h)')
        task_b = df_2026B.groupby(['책무명', '과업명'])['수행시간'].sum().rename('하반기(h)')
        task_df = pd.concat([task_a, task_b], axis=1).fillna(0).reset_index()
        task_df['증감(h)'] = task_df['하반기(h)'] - task_df['상반기(h)']
        task_df = task_df[task_df['증감(h)'].abs() > 0.5].sort_values('증감(h)')
        if task_df.empty:
            st.write("상반기 대비 변경된 과업이 없습니다.")
        else:
            st.dataframe(task_df.style.format({'상반기(h)': '{:,.0f}', '하반기(h)': '{:,.0f}', '증감(h)': '{:+,.0f}'}),
                         hide_index=True, use_container_width=True)

    st.info(f"""
🚀 **[Conclusion & Action Plan]**
* **현황:** 하반기에도 타이트한 실무 인원(실제 {CUR_PPL}명 vs 필요 {req_ppl[2]:.1f}명)으로 운영 중이며, 상반기({req_ppl[1]:.1f}명) 대비 필요 인원이 {req_ppl[2] - req_ppl[1]:+.1f}명 변동되었습니다.
* **주요 원인:** 기획부팀 내 CS지원팀 업무 일부가 추가되었고, 마케팅부팀의 용도별 담당자가 통합에서 분리 체계로 전환되었습니다.
* **극복 전략:** 방대한 분석/관리 업무를 데이터 분석 기반 파이썬 웹앱으로 자동화하여 극단적인 효율(Reduce)을 지속 달성합니다.
* **향후 집중 방향:** 시스템으로 확보한 여력을 용도별 전문성 강화와 **'마케팅분석 고도화 & 시각화 결과 공유'** 등에 투입하여 비즈니스 임팩트를 높이겠습니다.
""")

    st.divider()

    # =====================================================================
    # 분석 3: 책무별 업무량 정밀 분석 및 연도별 비교
    # =====================================================================
    st.subheader("📊 3. 책무별 업무량 정밀 분석 및 연도별 비교")

    st.markdown("2025년, 2026년 상반기, 2026년 하반기의 수행시간 증감 내역과, 현재(26년 하반기) 기준 직무/책무별 세부 수준을 확인하실 수 있습니다.")

    # ------------------ (1) 차트 렌더링 ------------------
    st.markdown("<br>**[ 2025년 vs 2026년 상반기 vs 2026년 하반기 책무별 수행시간 변동 비교 그래프 ]**", unsafe_allow_html=True)

    comp_df = diff_df.reset_index()
    plot_df = comp_df[(comp_df['2025년(h)'] > 0) | (comp_df['2026년 상반기(h)'] > 0) | (comp_df['2026년 하반기(h)'] > 0)]
    plot_df = plot_df.sort_values(by='2026년 하반기(h)', ascending=False)

    target_col = '업무(영업)용 마케팅'
    if target_col in plot_df['책무명'].values:
        target_row = plot_df[plot_df['책무명'] == target_col]
        plot_df = plot_df[plot_df['책무명'] != target_col]
        plot_df = pd.concat([plot_df, target_row])

    fig_comp = go.Figure()
    fig_comp.add_trace(go.Bar(x=plot_df['책무명'], y=plot_df['2025년(h)'], name=LBL_25, marker_color=COLOR_25))
    fig_comp.add_trace(go.Bar(x=plot_df['책무명'], y=plot_df['2026년 상반기(h)'], name=LBL_A, marker_color=COLOR_A))
    fig_comp.add_trace(go.Bar(x=plot_df['책무명'], y=plot_df['2026년 하반기(h)'], name=LBL_B, marker_color=COLOR_B))

    fig_comp.update_layout(
        barmode='group',
        height=500,
        xaxis_title="책무명",
        yaxis_title="연간 과업 수행시간 (Hours)",
        xaxis=dict(categoryorder='array', categoryarray=plot_df['책무명']),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    st.plotly_chart(fig_comp, use_container_width=True)

    # ------------------ (2) 표 렌더링 ------------------
    st.markdown("<br>**[ 2025년 vs 2026년 상반기 vs 2026년 하반기 직무/책무별 상세 업무량 분석 표 ]**", unsafe_allow_html=True)

    # 책무 → 직무 매핑 (하반기 > 상반기 > 2025 우선)
    job_map = (pd.concat([df_2026B[['책무명', '직무명']], df_2026A[['책무명', '직무명']], df_2025[['책무명', '직무명']]])
               .drop_duplicates(subset='책무명', keep='first'))

    # 하반기 기준 과업수/중요도/난이도/숙련도
    stat_26 = df_2026B.groupby('책무명').agg(
        과업수=('과업명', 'count'),
        중요도=('중요도', 'mean'),
        난이도=('난이도', 'mean'),
        숙련도=('숙련도', 'mean')
    ).reset_index()

    def get_mode(x):
        return x.mode().iloc[0] if not x.mode().empty else None

    grade_26 = df_2026B.groupby('책무명')['업무등급'].agg(get_mode).reset_index()

    agg = comp_df.rename(columns={
        '2025년(h)': '25년 수행시간',
        '2026년 상반기(h)': '26년 상반기 수행시간',
        '2026년 하반기(h)': '26년 하반기 수행시간',
        '증감(h)': '증감(h)(상반기→하반기)'
    })
    agg = (agg.merge(job_map, on='책무명', how='left')
              .merge(stat_26, on='책무명', how='left')
              .merge(grade_26, on='책무명', how='left'))

    agg['과업수'] = agg['과업수'].fillna(0)

    total_hours_26 = agg['26년 하반기 수행시간'].sum()
    agg['업무량 구성비(%)'] = (agg['26년 하반기 수행시간'] / total_hours_26 * 100).round(1)

    for c in ['중요도', '난이도', '숙련도']:
        agg[c] = agg[c].round(1)

    extra_jobs = [j for j in agg['직무명'].dropna().unique() if j not in custom_sort_order]
    agg['직무명'] = pd.Categorical(agg['직무명'], categories=custom_sort_order + extra_jobs, ordered=True)

    cols = ['직무명', '책무명', '과업수', '25년 수행시간', '26년 상반기 수행시간', '26년 하반기 수행시간',
            '증감(h)(상반기→하반기)', '업무량 구성비(%)', '중요도', '난이도', '숙련도', '업무등급']

    agg = agg[cols].sort_values(['직무명'], kind='stable')

    total_row = pd.DataFrame([{
        '직무명': '합계',
        '책무명': '-',
        '과업수': agg['과업수'].sum(),
        '25년 수행시간': agg['25년 수행시간'].sum(),
        '26년 상반기 수행시간': agg['26년 상반기 수행시간'].sum(),
        '26년 하반기 수행시간': agg['26년 하반기 수행시간'].sum(),
        '증감(h)(상반기→하반기)': agg['증감(h)(상반기→하반기)'].sum(),
        '업무량 구성비(%)': 100.0,
        '중요도': None,
        '난이도': None,
        '숙련도': None,
        '업무등급': '-'
    }])
    agg['직무명'] = agg['직무명'].astype(str)
    agg = pd.concat([agg, total_row], ignore_index=True)

    def color_negative_red(val):
        if pd.isna(val): return ''
        try:
            val_num = float(val)
            color = 'red' if val_num < 0 else 'blue' if val_num > 0 else 'black'
            return f'color: {color}'
        except ValueError:
            return ''

    def highlight_total(row):
        if row['직무명'] == '합계':
            return ['background-color: #e6e6e6; font-weight: bold'] * len(row)
        return [''] * len(row)

    st.dataframe(agg.style.format({
        '과업수': '{:.0f}',
        '25년 수행시간': '{:,.0f}', '26년 상반기 수행시간': '{:,.0f}', '26년 하반기 수행시간': '{:,.0f}',
        '증감(h)(상반기→하반기)': '{:,.0f}',
        '업무량 구성비(%)': '{:.1f}',
        '중요도': '{:.1f}', '난이도': '{:.1f}', '숙련도': '{:.1f}'
    }, na_rep="-").map(color_negative_red, subset=['증감(h)(상반기→하반기)']).apply(highlight_total, axis=1),
        hide_index=True, use_container_width=True)

    st.divider()

    # =====================================================================
    # 분석 4: 기획부팀 / 마케팅부팀 분리 (직무 기준)
    #   기획부팀   : 영업기획/관리, 일반관리
    #   마케팅부팀 : 주택용 수요개발, 주택용 외 수요개발
    #   ※ 팀리더는 팀 공통이므로 이 섹션에서는 제외
    # =====================================================================
    st.subheader("🏢 4. 기획부팀 · 마케팅부팀 분리 분석 (직무 기준)")
    st.markdown(
        "팀리더를 제외하고, 직무를 기준으로 **기획부팀**(영업기획/관리, 일반관리)과 **마케팅부팀**(주택용 수요개발, 주택용 외 수요개발)의 "
        "과업시간과 적정 필요 인원을 나누어 비교합니다."
    )

    TEAM_JOBS = {
        '기획부팀': ['영업기획/관리', '일반관리'],
        '마케팅부팀': ['주택용 수요개발', '주택용 외 수요개발'],
    }
    TEAM_ICON = {'기획부팀': '🗂️', '마케팅부팀': '📣'}

    def team_hours(d, jobs):
        return d[d['직무명'].isin(jobs)]['수행시간'].sum()

    # ---------- (1) 팀별 요약표 ----------
    summary_rows = []
    for team, jobs in TEAM_JOBS.items():
        hrs = [team_hours(d, jobs) for d in dfs]
        summary_rows.append({
            '팀': team,
            '포함 직무': ', '.join(jobs),
            f'{LBL_25} 수행시간(h)': hrs[0],
            f'{LBL_A} 수행시간(h)': hrs[1],
            f'{LBL_B} 수행시간(h)': hrs[2],
            '증감(h)(상반기→하반기)': hrs[2] - hrs[1],
            f'{LBL_A} 필요 인원(명)': hrs[1] / std_hours,
            f'{LBL_B} 필요 인원(명)': hrs[2] / std_hours,
            '필요 인원 증감(명)': (hrs[2] - hrs[1]) / std_hours,
        })
    summary_df = pd.DataFrame(summary_rows)

    st.markdown("**[ 팀별 과업시간 및 필요 인원 요약 ]**")
    st.dataframe(
        summary_df.style.format({
            f'{LBL_25} 수행시간(h)': '{:,.0f}',
            f'{LBL_A} 수행시간(h)': '{:,.0f}',
            f'{LBL_B} 수행시간(h)': '{:,.0f}',
            '증감(h)(상반기→하반기)': '{:+,.0f}',
            f'{LBL_A} 필요 인원(명)': '{:.1f}',
            f'{LBL_B} 필요 인원(명)': '{:.1f}',
            '필요 인원 증감(명)': '{:+.1f}',
        }).map(color_negative_red, subset=['증감(h)(상반기→하반기)', '필요 인원 증감(명)']),
        hide_index=True, use_container_width=True
    )

    # ---------- (2) 팀별 상세 (좌: 기획부팀 / 우: 마케팅부팀) ----------
    def duty_table(jobs):
        parts = [
            d[d['직무명'].isin(jobs)].groupby(['직무명', '책무명'])['수행시간'].sum().rename(lbl)
            for d, lbl in zip(dfs, labels)
        ]
        t = pd.concat(parts, axis=1).fillna(0).reset_index()
        t['증감(h)(상반기→하반기)'] = t[LBL_B] - t[LBL_A]
        t['직무명'] = pd.Categorical(t['직무명'], categories=jobs, ordered=True)
        t = t.sort_values(['직무명', LBL_B], ascending=[True, False], kind='stable')
        t['직무명'] = t['직무명'].astype(str)
        total = {'직무명': '합계', '책무명': '-'}
        for c in [LBL_25, LBL_A, LBL_B, '증감(h)(상반기→하반기)']:
            total[c] = t[c].sum()
        return pd.concat([t, pd.DataFrame([total])], ignore_index=True)

    def highlight_total_row(row):
        if row['직무명'] == '합계':
            return ['background-color: #e6e6e6; font-weight: bold'] * len(row)
        return [''] * len(row)

    team_cols = st.columns(2)
    for col, (team, jobs) in zip(team_cols, TEAM_JOBS.items()):
        with col:
            st.markdown(f"#### {TEAM_ICON[team]} {team}")
            st.caption("담당 직무: " + " / ".join(jobs))

            hrs = [team_hours(d, jobs) for d in dfs]
            m1, m2, m3 = st.columns(3)
            m1.metric(f"{LBL_B} 수행시간", f"{hrs[2]:,.0f}h",
                      delta=f"{hrs[2] - hrs[1]:+,.0f}h (상반기 대비)", delta_color="off")
            m2.metric(f"{LBL_B} 필요 인원", f"{hrs[2] / std_hours:.1f}명",
                      delta=f"{(hrs[2] - hrs[1]) / std_hours:+.1f}명 (상반기 대비)", delta_color="off")
            m3.metric(f"{LBL_A} 필요 인원", f"{hrs[1] / std_hours:.1f}명")

            # 직무별 누적 막대
            fig_team = go.Figure()
            for job in jobs:
                vals = [d[d['직무명'] == job]['수행시간'].sum() for d in dfs]
                fig_team.add_trace(go.Bar(
                    name=job, x=labels, y=vals, marker_color=JOB_COLORS.get(job, '#64748B'),
                    marker_line=dict(color="white", width=1.5),
                    text=[f"{v:,.0f}h" if v > 0 else "" for v in vals],
                    textposition="inside", insidetextanchor="middle",
                    textfont=dict(size=12, color="white"),
                    hovertemplate="%{x}<br>" + job + ": %{y:,.0f}h<extra></extra>"
                ))
            fig_team.add_trace(go.Scatter(
                x=labels, y=hrs, mode="text", showlegend=False, hoverinfo="skip",
                text=[f"<b>{v:,.0f}h</b>" for v in hrs], textposition="top center",
                textfont=dict(size=14, color="#0F172A")
            ))
            style_fig(fig_team, height=450)
            fig_team.update_layout(barmode="stack", bargap=0.35, margin=dict(t=70, b=50, l=60, r=30))
            fig_team.update_yaxes(range=[0, max(hrs + [1]) * 1.15], tickformat=",", title_text="연간 수행시간 (h)")
            st.plotly_chart(fig_team, use_container_width=True, config=PLOT_CFG)

            # 책무별 상세표
            t = duty_table(jobs)
            st.dataframe(
                t.style.format({
                    LBL_25: '{:,.0f}', LBL_A: '{:,.0f}', LBL_B: '{:,.0f}',
                    '증감(h)(상반기→하반기)': '{:+,.0f}',
                }).map(color_negative_red, subset=['증감(h)(상반기→하반기)']).apply(highlight_total_row, axis=1),
                hide_index=True, use_container_width=True
            )

else:
    st.warning("지정된 엑셀 파일을 찾을 수 없습니다. 파일명에 _2025B_ / _2026A_ / _2026B_ 가 포함되어 있는지 확인해주세요.")
    missing = [k for k, v in [('_2025B_ (2025년)', file_2025), ('_2026A_ (2026년 상반기)', file_2026A), ('_2026B_ (2026년 하반기)', file_2026B)] if not v]
    st.write("찾지 못한 파일:", missing)
    st.write("현재 폴더의 파일 목록:", sorted(os.listdir(BASE_DIR)))
