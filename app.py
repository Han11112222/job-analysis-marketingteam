st.divider()

    # =====================================================================
    # 분석 4: 기획부팀 / 마케팅부팀 분리 (직무 기준)
    #   기획부팀   : 영업기획/관리
    #   마케팅부팀 : 주택용 수요개발, 주택용 외 수요개발
    #   ※ 팀리더, 일반관리는 팀 공통이므로 이 섹션에서는 제외
    # =====================================================================
    st.subheader("🏢 4. 기획부팀 · 마케팅부팀 분리 분석 (직무 기준)")
    st.markdown(
        "팀 공통 직무(팀리더, 일반관리)를 제외하고, 직무를 기준으로 **기획부팀**과 **마케팅부팀**의 "
        "과업시간과 적정 필요 인원을 나누어 비교합니다."
    )

    TEAM_JOBS = {
        '기획부팀': ['영업기획/관리'],
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
