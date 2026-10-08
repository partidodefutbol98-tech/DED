import streamlit as st
import pandas as pd
import datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# 1. 페이지 기본 설정
st.set_page_config(page_title="대구FC 선수 평가 대시보드", layout="wide")

# 2. 대시보드 커스텀 CSS (하단 차트 확대 및 A4 1페이지 인쇄 최적화)
st.markdown("""
    <style>
    @import url("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.css");

    .main {
        background-color: #f8f9fa;
        font-family: "Pretendard", -apple-system, BlinkMacSystemFont, system-ui, Roboto, sans-serif !important;
    }

    /* 카드 형태 박스 스타일 */
    .card-box {
        background-color: white;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.02);
        border: 1px solid #f1f5f9;
        margin-bottom: 10px;
    }

    /* 프로필 텍스트 라벨 및 수치 크기 */
    .profile-name {
        font-size: 22px !important;
        font-weight: 800 !important;
        color: #0f172a !important;
        margin-bottom: 4px !important;
    }
    
    .profile-info {
        font-size: 14px !important;
        color: #334155 !important;
        line-height: 1.6 !important;
    }

    .salary-title {
        font-size: 13px !important;
        font-weight: 700 !important;
        color: #0f172a !important;
    }

    .salary-value {
        font-size: 15px !important;
        font-weight: 800 !important;
        color: #0284c7 !important;
    }

    /* 계약 옵션 강조 박스 */
    .opt-box {
        background-color: #e0f2fe;
        border-left: 4px solid #0284c7;
        padding: 6px 10px;
        border-radius: 4px;
        font-size: 12px !important;
        color: #0369a1;
        margin-top: 3px;
        margin-bottom: 6px;
        line-height: 1.4;
    }

    /* 중앙 파란색 코멘트 배너 */
    .comment-banner {
        background-color: #3b62be;
        color: white;
        text-align: center;
        font-size: 16px !important;
        font-weight: 700;
        line-height: 1.5;
        padding: 12px 18px;
        border-radius: 8px;
        margin: 10px 0px;
        box-shadow: 0px 4px 10px rgba(59, 98, 190, 0.15);
        word-break: keep-all;
    }

    /* 평점 숫자 서식 */
    .metric-value {
        font-size: 24px !important;
        font-weight: 800 !important;
        color: #1e293b;
        text-align: center;
    }
    .metric-label {
        font-size: 12px !important;
        font-weight: 600;
        color: #64748b;
        text-align: center;
        margin-bottom: 2px;
    }

    /* A4 용지 가로 1페이지 인쇄(Ctrl+P) 맞춤 스타일 */
    @media print {
        @page {
            size: A4 landscape;
            margin: 4mm;
        }
        body {
            zoom: 78%;
        }
        section[data-testid="stSidebar"],
        header, footer, .stDeployButton {
            display: none !important;
        }
        .main .block-container {
            max-width: 100% !important;
            padding: 0 !important;
        }
        .card-box {
            box-shadow: none !important;
            border: 1px solid #cbd5e1 !important;
            padding: 8px 12px !important;
            margin-bottom: 6px !important;
            page-break-inside: avoid;
        }
        .comment-banner {
            padding: 8px 12px !important;
            margin: 6px 0px !important;
        }
    }

    /* 사이드바 애니메이션 */
    @keyframes fadeInUp {
        0% { opacity: 0; transform: translateY(12px); }
        100% { opacity: 1; transform: translateY(0); }
    }

    .rank-card {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 12px 14px;
        border: 1px solid #e2e8f0;
        margin-top: 14px;
        box-shadow: 0px 2px 6px rgba(0,0,0,0.02);
    }

    .rank-item {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 6px 0;
        border-bottom: 1px solid #f8fafc;
        animation: fadeInUp 0.4s ease-out forwards;
        opacity: 0;
    }

    .rank-item:nth-child(1) { animation-delay: 0.05s; }
    .rank-item:nth-child(2) { animation-delay: 0.10s; }
    .rank-item:nth-child(3) { animation-delay: 0.15s; }
    .rank-item:nth-child(4) { animation-delay: 0.20s; }
    .rank-item:nth-child(5) { animation-delay: 0.25s; }
    .rank-item:nth-child(6) { animation-delay: 0.30s; }
    .rank-item:nth-child(7) { animation-delay: 0.35s; }
    .rank-item:nth-child(8) { animation-delay: 0.40s; }
    .rank-item:nth-child(9) { animation-delay: 0.45s; }
    .rank-item:nth-child(10) { animation-delay: 0.50s; }
    </style>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# [데이터 로드] 엑셀 #1 ~ #5 시트 통합 읽기
# ----------------------------------------------------
EXCEL_FILE = '선수 평가 대시보드 rep..xlsx'

@st.cache_data(ttl=2)
def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheets = xls.sheet_names
    
    df_profile = pd.read_excel(xls, sheet_name='#1')
    df_matches = pd.read_excel(xls, sheet_name='#2')
    df_k2_rank = pd.read_excel(xls, sheet_name='#3') if '#3' in sheets else pd.DataFrame()
    df_radar   = pd.read_excel(xls, sheet_name='#4') if '#4' in sheets else pd.DataFrame()
    df_trends  = pd.read_excel(xls, sheet_name='#5') if '#5' in sheets else pd.DataFrame()
    
    df_profile = df_profile.dropna(subset=['선수 이름']).copy()
    df_profile['선수 이름'] = df_profile['선수 이름'].astype(str).str.strip()
    df_matches['선수이름'] = df_matches['선수이름'].astype(str).str.strip()
    
    if not df_k2_rank.empty and '선수이름' in df_k2_rank.columns:
        df_k2_rank['선수이름'] = df_k2_rank['선수이름'].astype(str).str.strip()
    if not df_radar.empty and '선수이름' in df_radar.columns:
        df_radar['선수이름'] = df_radar['선수이름'].astype(str).str.strip()
    if not df_trends.empty and '선수이름' in df_trends.columns:
        df_trends['선수이름'] = df_trends['선수이름'].astype(str).str.strip()
        
    return df_profile, df_matches, df_k2_rank, df_radar, df_trends

try:
    df_profile, df_matches, df_k2_rank, df_radar, df_trends = load_data(EXCEL_FILE)
except Exception as e:
    st.error(f"엑셀 파일({EXCEL_FILE})을 읽는 중 오류가 발생했습니다: {e}")
    st.stop()

# 연봉 자동 포맷 변환 함수
def format_salary(val):
    try:
        if pd.isna(val) or val == '-' or str(val).strip() == '':
            return "-"
        val = float(val)
        if val == 0:
            return "-"
        if val < 10000000:
            return f"${val:,.0f}"
        if val >= 100000000:
            uk = int(val // 100000000)
            man = int((val % 100000000) // 10000)
            if man > 0:
                return f"{uk}억 {man:,}만원"
            return f"{uk}억원"
        else:
            man = int(val // 10000)
            return f"{man:,}만원"
    except:
        return str(val) if pd.notna(val) else "-"

# ----------------------------------------------------
# [사이드바] 선수 선택 및 실시간 애니메이션 순위
# ----------------------------------------------------
st.sidebar.title("대구FC 선수 관리")
player_list = df_profile['선수 이름'].unique().tolist()
default_index = player_list.index("고동민") if "고동민" in player_list else 0
selected_name = st.sidebar.selectbox("선수를 선택하세요:", player_list, index=default_index)

p = df_profile[df_profile['선수 이름'] == selected_name].iloc[0]
p_matches = df_matches[df_matches['선수이름'] == selected_name].copy()

# 실시간 순위 애니메이션 적용 Top 10 위젯
top10_dept = df_profile.dropna(subset=['전력강화평점']).sort_values('전력강화평점', ascending=False).head(10).reset_index(drop=True)

rank_html_items = ""
for idx, row in top10_dept.iterrows():
    rank = idx + 1
    name = row['선수 이름']
    score = row['전력강화평점']
    
    if rank == 1:
        badge_style = "background-color: #0284c7; color: white;"
    elif rank == 2:
        badge_style = "background-color: #0369a1; color: white;"
    elif rank == 3:
        badge_style = "background-color: #0f172a; color: white;"
    else:
        badge_style = "background-color: #f1f5f9; color: #64748b;"
        
    rank_html_items += f"""<div class="rank-item">
<div style="display:flex; align-items:center;">
<span style="display:inline-block; width:20px; height:20px; line-height:20px; text-align:center; border-radius:4px; font-weight:700; font-size:11px; margin-right:8px; {badge_style}">{rank}</span>
<span style="font-size:13px; font-weight:600; color:#1e293b;">{name}</span>
</div>
<span style="font-size:12px; font-weight:700; color:#0284c7;">{score:.2f}점</span>
</div>"""

rank_card_html = f"""<div class="rank-card">
<div style="font-size:13px; font-weight:800; color:#0f172a; margin-bottom:8px; border-bottom:2px solid #0284c7; padding-bottom:6px;">
전력강화부 평점 Top 10
</div>
{rank_html_items}
</div>"""

st.sidebar.markdown(rank_card_html, unsafe_allow_html=True)

# ----------------------------------------------------
# [상단 헤더]
# ----------------------------------------------------
today_date = datetime.datetime.now().strftime("%Y.%m.%d")
col_head_left, col_head_right = st.columns([7, 3])

with col_head_left:
    st.markdown("""
        <div style='display: flex; align-items: baseline; gap: 12px; padding-top: 2px;'>
            <span style='font-size: 24px; font-weight: 800; color: #000000;'>대구FC 선수 평가 대시보드</span>
            <span style='font-size: 13px; font-weight: 600; color: #475569;'>DaeguFC Player Evaluation Dashboard</span>
        </div>
    """, unsafe_allow_html=True)

with col_head_right:
    st.markdown(f"""
        <div style='text-align: right; font-size: 12px; color: #475569; font-weight: 500; padding-top: 8px;'>
            기준일: {today_date} / 전력강화팀 여지민
        </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 1px; background-color: #a3c1e0; margin-top: 4px; margin-bottom: 12px;'></div>", unsafe_allow_html=True)

# ----------------------------------------------------
# [상단 구역] 프로필/증명사진/연봉 vs 출전시간/평점
# ----------------------------------------------------
top_left, top_right = st.columns([1.15, 1])

with top_left:
    st.markdown("<div class='card-box'>", unsafe_allow_html=True)
    c_prof, c_sal = st.columns([1.3, 1])
    
    with c_prof:
        img_col, info_col = st.columns([1, 2.2])
        with img_col:
            st.image("https://via.placeholder.com/100x125/0284c7/ffffff?text=PROFILE", use_container_width=True)
            
        with info_col:
            back_num = f"No. {int(p['번호'])}" if pd.notna(p['번호']) else "No. -"
            st.markdown(f"<div class='profile-name'>{selected_name} ({back_num})</div>", unsafe_allow_html=True)
            birth_str = str(p['생년월일'])[:10] if pd.notna(p['생년월일']) else "-"
            st.markdown(f"""
                <div class='profile-info'>
                    • <b>생년월일:</b> {birth_str}<br>
                    • <b>신체조건:</b> {p['신제 조건']}<br>
                    • <b>계약기간:</b> {p['계약기간']}
                </div>
            """, unsafe_allow_html=True)
            
    with c_sal:
        sal_2026 = format_salary(p['연봉_2026'])
        opt_2026 = str(p['옵션_2026']) if pd.notna(p['옵션_2026']) and str(p['옵션_2026']) != '-' else "없음"
        st.markdown(f"<span class='salary-title'>2026 연봉:</span> <span class='salary-value'>{sal_2026}</span>", unsafe_allow_html=True)
        st.markdown(f"<div class='opt-box'><b>2026시즌 계약 옵션사항:</b><br>{opt_2026}</div>", unsafe_allow_html=True)
        
        sal_2025 = format_salary(p['연봉_2025'])
        opt_2025 = str(p['옵션_2025']) if pd.notna(p['옵션_2025']) and str(p['옵션_2025']) != '-' else "없음"
        st.markdown(f"<span class='salary-title'>2025 연봉:</span> <span style='font-size:14px; font-weight:700; color:#64748b;'>{sal_2025}</span>", unsafe_allow_html=True)
        st.markdown(f"<div class='opt-box'><b>2025시즌 계약 옵션사항:</b><br>{opt_2025}</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

with top_right:
    st.markdown("<div class='card-box'>", unsafe_allow_html=True)
    st.markdown(f"<b style='font-size:13px; color:#0f172a;'>라운드별 경기 결과 및 출전시간 ({selected_name})</b>", unsafe_allow_html=True)
    
    if not p_matches.empty:
        match_rounds = p_matches['라운드'].tolist()
        match_minutes = p_matches['출전시간'].fillna(0).tolist()
        match_results = p_matches['경기결과'].fillna('무').tolist()
        
        fig_match = make_subplots(specs=[[{"secondary_y": True}]])
        fig_match.add_trace(go.Bar(x=match_rounds, y=match_minutes, name="출전시간(분)", marker_color="#bae6fd", opacity=0.8), secondary_y=False)
        
        color_map = {'승': '#22c55e', '무': '#64748b', '패': '#ef4444'}
        y_map = {'승': 3, '무': 2, '패': 1}
        
        for res in ['승', '무', '패']:
            r_x = [match_rounds[i] for i in range(len(match_rounds)) if match_results[i] == res]
            r_y = [y_map[res] for _ in r_x]
            if r_x:
                fig_match.add_trace(go.Scatter(x=r_x, y=r_y, mode='markers', name=res, marker=dict(color=color_map[res], size=5)), secondary_y=True)
            
        fig_match.update_layout(
            font=dict(size=9.5),
            height=110, 
            margin=dict(l=5, r=5, t=5, b=5), 
            template="plotly_white",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        fig_match.update_yaxes(title_text="출전시간", secondary_y=False, range=[0, 105])
        fig_match.update_yaxes(showticklabels=False, secondary_y=True, range=[0, 4])
        st.plotly_chart(fig_match, use_container_width=True)
    else:
        st.info("해당 선수의 출전 기록 데이터가 없습니다.")
        
    st.markdown("<div style='height:1px; background-color:#e2e8f0; margin:4px 0px;'></div>", unsafe_allow_html=True)
    
    m1, m2, m3 = st.columns(3)
    director_score = f"{p['단장 평점']:.2f}" if pd.notna(p['단장 평점']) else "-"
    dept_score = f"{p['전력강화평점']:.2f}" if pd.notna(p['전력강화평점']) else "-"
    league_score = f"{p['연맹평점']:.2f}" if pd.notna(p['연맹평점']) else "-"
    
    with m1:
        st.markdown(f"<div class='metric-label'>단장님 평가</div><div class='metric-value'>{director_score}</div>", unsafe_allow_html=True)
    with m2:
        st.markdown(f"<div class='metric-label'>전력강화부</div><div class='metric-value'>{dept_score}</div>", unsafe_allow_html=True)
    with m3:
        st.markdown(f"<div class='metric-label'>연맹 평점</div><div class='metric-value'>{league_score}</div>", unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------
# [중앙 구역] 코멘트 배너
# ----------------------------------------------------
player_comment = None
for col in p.index:
    col_str = str(col).replace(" ", "").lower()
    if any(k in col_str for k in ['코멘트', '총평', '비고', '평가내용', '상세평가', '특이사항', '강화부']):
        val = str(p[col]).strip()
        if pd.notna(p[col]) and val not in ['', '-', 'nan', 'None']:
            player_comment = val
            break

if not player_comment:
    player_comment = "선수에 대해 작성된 코멘트가 없습니다."

st.markdown(f"<div class='comment-banner'>{player_comment}</div>", unsafe_allow_html=True)

# ----------------------------------------------------
# [하단 구역] 3분할 가로 정렬 리포트 (#3, #4, #5) - 높이 상향 확대 (300px)
# ----------------------------------------------------
col_report1, col_report2, col_report3 = st.columns(3)

# ----------------------------------------------------
# 1. #3 시트: K2 동일 포지션 선수 순위
# ----------------------------------------------------
with col_report1:
    st.markdown("<div class='card-box'>", unsafe_allow_html=True)
    st.markdown("<div style='display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;'>"
                "<span style='font-weight:700; font-size:15px; color:#0f172a;'>K2 동일 포지션 선수 순위</span>"
                "<span style='color:#ef4444; font-weight:600; font-size:11px;'>* 출전경기 20경기 이상 / 출전시간 1200분 이상</span>"
                "</div>", unsafe_allow_html=True)
    
    p_pos_row = df_k2_rank[df_k2_rank['선수이름'] == selected_name] if not df_k2_rank.empty else pd.DataFrame()
    p_pos = p_pos_row['포지션'].values[0] if not p_pos_row.empty else None
    
    if p_pos and p_pos != '-':
        df_pos = df_k2_rank[df_k2_rank['포지션'] == p_pos].copy()
        df_pos['PEI_num'] = pd.to_numeric(df_pos['PEI 종합 점수'], errors='coerce')
        df_pos = df_pos.dropna(subset=['PEI_num']).sort_values('PEI_num', ascending=False).reset_index(drop=True)
        
        top10 = df_pos.head(10).copy()
        
        if selected_name in top10['선수이름'].values:
            chart_df = top10.copy()
        else:
            selected_row = df_pos[df_pos['선수이름'] == selected_name]
            if not selected_row.empty:
                chart_df = pd.concat([top10, selected_row]).reset_index(drop=True)
            else:
                chart_df = top10.copy()
                
        chart_df_plot = chart_df.iloc[::-1]
        colors = ['#0284c7' if x == selected_name else '#e2e8f0' for x in chart_df_plot['선수이름']]
        
        fig_bar = go.Figure(go.Bar(
            x=chart_df_plot['PEI_num'],
            y=chart_df_plot['선수이름'],
            orientation='h',
            text=chart_df_plot['PEI_num'].round(2),
            textposition='outside',
            marker_color=colors
        ))
        fig_bar.update_layout(
            font=dict(size=10),
            height=300,
            margin=dict(l=75, r=30, t=10, b=10),
            template="plotly_white",
            xaxis=dict(range=[0, max(chart_df_plot['PEI_num'].max() * 1.15, 100)])
        )
        st.plotly_chart(fig_bar, use_container_width=True)
    else:
        st.info("포지션 순위 데이터가 없습니다.")
        
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------
# 2. #4 시트: 세부 지표 비교
# ----------------------------------------------------
with col_report2:
    st.markdown("<div class='card-box'>", unsafe_allow_html=True)
    st.markdown("<div style='display:flex; align-items:center; justify-content:space-between; margin-bottom:8px;'>"
                "<span style='font-weight:700; font-size:15px; color:#0f172a;'>세부 지표 비교</span>"
                "<span style='color:#ef4444; font-weight:600; font-size:11px;'>* 출전경기 20경기 이상 / 출전시간 1200분 이상</span>"
                "</div>", unsafe_allow_html=True)
    
    p_radar = df_radar[df_radar['선수이름'] == selected_name] if not df_radar.empty else pd.DataFrame()
    
    if not p_radar.empty and p_pos:
        same_pos_radar = df_radar[df_radar['포지션'] == p_pos].copy()
        same_pos_radar['PEI_num'] = pd.to_numeric(same_pos_radar['PEI 종합 점수'], errors='coerce')
        top1_radar_row = same_pos_radar.sort_values('PEI_num', ascending=False).iloc[0]
        
        row_sel = p_radar.iloc[0]
        cats, sel_vals, top_vals = [], [], []
        
        for i in range(1, 9):
            cn_matches = [c for c in df_radar.columns if c.strip() == f'세부지표_{i}']
            cs_matches = [c for c in df_radar.columns if c.strip() == f'세부지표_{i} 점수']
            
            if cn_matches and cs_matches:
                cn, cs = cn_matches[0], cs_matches[0]
                name_v = row_sel[cn]
                score_sel = row_sel[cs]
                score_top = top1_radar_row[cs]
                
                if pd.notna(name_v) and str(name_v).strip() not in ['-', '']:
                    cats.append(str(name_v).strip())
                    try: sel_vals.append(float(score_sel))
                    except: sel_vals.append(0.0)
                    try: top_vals.append(float(score_top))
                    except: top_vals.append(0.0)
                    
        if cats:
            top1_name = top1_radar_row['선수이름']
            fig_radar = go.Figure()
            
            fig_radar.add_trace(go.Scatterpolar(
                r=top_vals + [top_vals[0]],
                theta=cats + [cats[0]],
                fill='toself',
                name=f'1위 ({top1_name})',
                line=dict(color='#94a3b8', dash='dash'),
                fillcolor='rgba(148, 163, 184, 0.15)'
            ))
            
            fig_radar.add_trace(go.Scatterpolar(
                r=sel_vals + [sel_vals[0]],
                theta=cats + [cats[0]],
                fill='toself',
                name=selected_name,
                line=dict(color='#0284c7', width=2),
                fillcolor='rgba(2, 132, 199, 0.3)'
            ))
            
            fig_radar.update_layout(
                polar=dict(
                    radialaxis=dict(visible=True, range=[0, 100], tickfont=dict(size=7)),
                    angularaxis=dict(tickfont=dict(size=8.5))
                ),
                font=dict(size=9.5),
                height=300,
                margin=dict(l=35, r=35, t=15, b=15),
                template="plotly_white",
                legend=dict(orientation="h", yanchor="bottom", y=-0.18, xanchor="center", x=0.5)
            )
            st.plotly_chart(fig_radar, use_container_width=True)
        else:
            st.info("세부 지표 데이터가 없습니다.")
    else:
        st.info("#4 시트 데이터가 없거나 선수 정보를 찾을 수 없습니다.")
        
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------
# 3. #5 시트: 2025 vs 2026 PEI 비교
# ----------------------------------------------------
with col_report3:
    st.markdown("<div class='card-box'>", unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom:8px;'><span style='font-weight:700; font-size:15px; color:#0f172a;'>2025 vs 2026 PEI 비교</span></div>", unsafe_allow_html=True)
    
    p_trends = df_trends[df_trends['선수이름'] == selected_name].dropna(subset=['PEI 카테고리 및 종합 점수']) if not df_trends.empty else pd.DataFrame()
    
    if not p_trends.empty:
        df_display = p_trends.copy()
        
        col_2025 = 2025 if 2025 in df_display.columns else '2025'
        col_2026 = 2026 if 2026 in df_display.columns else '2026'
        
        def calc_diff(row):
            v25 = pd.to_numeric(row[col_2025], errors='coerce')
            v26 = pd.to_numeric(row[col_2026], errors='coerce')
            if pd.notna(v25) and pd.notna(v26):
                diff = v26 - v25
                if diff > 0:
                    return f"▲ {abs(diff):.2f}"
                elif diff < 0:
                    return f"▼ {abs(diff):.2f}"
                else:
                    return "0.00"
            return "-"

        df_display['Diff.'] = df_display.apply(calc_diff, axis=1)
        
        show_df = df_display[['PEI 카테고리 및 종합 점수', col_2025, col_2026, 'Diff.' ]].copy()
        show_df.columns = ['PEI 카테고리 및 종합 점수', '2025', '2026', 'Diff.']
        
        show_df['2025'] = pd.to_numeric(show_df['2025'], errors='coerce').map(lambda x: f"{x:.2f}" if pd.notna(x) else "-")
        show_df['2026'] = pd.to_numeric(show_df['2026'], errors='coerce').map(lambda x: f"{x:.2f}" if pd.notna(x) else "-")
        
        def style_diff(val):
            if '▲' in str(val):
                return 'color: #22c55e; font-weight: bold;'
            elif '▼' in str(val):
                return 'color: #ef4444; font-weight: bold;'
            return 'color: #64748b;'

        try:
            styled_table = show_df.style.applymap(style_diff, subset=['Diff.'])
        except AttributeError:
            styled_table = show_df.style.map(style_diff, subset=['Diff.'])

        st.dataframe(styled_table, use_container_width=True, hide_index=True, height=300)
    else:
        st.info("#5 시트 비교 데이터가 없습니다.")
        
    st.markdown("</div>", unsafe_allow_html=True)