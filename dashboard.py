import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime
from collections import Counter
import re
import warnings
warnings.filterwarnings('ignore')

# 페이지 설정
st.set_page_config(
    page_title="2025 성장지원 워크샵 대시보드",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 커스텀 CSS - SK 브랜드 컬러 반영 및 필터 스타일 수정
st.markdown("""
<style>
    .main {
        padding: 0rem 1rem;
    }
    .stMetric {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .stMetric:hover {
        box-shadow: 0 4px 8px rgba(0,0,0,0.15);
        transition: all 0.3s;
    }
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .sidebar-content {
        background-color: #f5f5f5;
    }
    h1 {
        color: #ea002c;
    }
    h2 {
        color: #ff5800;
        border-bottom: 2px solid #ff5800;
        padding-bottom: 10px;
    }
    h3 {
        color: #333;
    }
    .stSelectbox label, .stMultiselect label {
        color: #ea002c;
        font-weight: bold;
    }
    /* 필터 옵션 스타일 수정 */
    .stMultiselect > div > div > div {
        font-size: 12px !important;
        background-color: transparent !important;
    }
    .stMultiselect [data-baseweb="tag"] {
        font-size: 12px !important;
        background-color: transparent !important;
    }
    /* 탭 스타일 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 8px;
        padding-left: 20px;
        padding-right: 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ea002c;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# 데이터 로드 함수
@st.cache_data
def load_data():
    """엑셀 파일에서 데이터를 로드합니다."""
    try:
        import os
        import sys
        
        # Streamlit Cloud 환경 확인
        current_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 파일명 (영문으로 변경 권장)
        file_names = [
            'dashboard_template.xlsx',  # 영문 파일명 (권장)
            '대시보드 템플릿.xlsx',  # 한글 파일명 (기존)
        ]
        
        file_path = None
        for filename in file_names:
            # 여러 경로 시도
            paths_to_try = [
                filename,  # 현재 디렉토리
                os.path.join(current_dir, filename),  # 스크립트 디렉토리
                os.path.join(os.getcwd(), filename),  # 작업 디렉토리
            ]
            
            for path in paths_to_try:
                if os.path.exists(path) and os.path.isfile(path):
                    file_path = path
                    break
            
            if file_path:
                break
        
        if file_path is None:
            # 디버깅 정보 표시
            st.error("⚠️ 엑셀 파일을 찾을 수 없습니다.")
            st.info("💡 다음 파일명 중 하나가 GitHub 저장소에 있는지 확인하세요:")
            for name in file_names:
                st.info(f"   - {name}")
            
            # 현재 디렉토리 파일 목록 표시 (디버깅용)
            with st.expander("🔍 디버깅 정보 (클릭하여 확인)"):
                st.write(f"현재 디렉토리: {os.getcwd()}")
                st.write(f"스크립트 디렉토리: {current_dir if 'current_dir' in locals() else 'N/A'}")
                st.write("디렉토리 내 파일 목록:")
                try:
                    files = os.listdir(os.getcwd())
                    for f in files:
                        if f.endswith('.xlsx'):
                            st.write(f"   ✓ {f}")
                        else:
                            st.write(f"   - {f}")
                except Exception as e:
                    st.write(f"   오류: {str(e)}")
            
            return None
        
        # 파일 읽기
        with st.spinner(f'📊 데이터 로드 중... ({os.path.basename(file_path)})'):
            # 각 시트를 DataFrame으로 읽기
            program_info = pd.read_excel(file_path, sheet_name='Program_Info')
            learners = pd.read_excel(file_path, sheet_name='Learners')
            certification = pd.read_excel(file_path, sheet_name='Certification')
            budget = pd.read_excel(file_path, sheet_name='Budget')
            instructors = pd.read_excel(file_path, sheet_name='Instructors')
            survey = pd.read_excel(file_path, sheet_name='Survey')
        
        # 날짜 형식 변환
        program_info['program_month'] = pd.to_datetime(program_info['program_month'])
        
        # 예산 계산 추가
        budget['actual_budget'] = budget['total_budget']
        
        # 직접비 총액 계산
        for idx, row in budget.iterrows():
            prog = program_info[program_info['program_id'] == row['program_id']].iloc[0]
            budget.loc[idx, 'total_direct_cost'] = row['direct_cost'] * prog['num_learners']
        
        return {
            'program_info': program_info,
            'learners': learners,
            'certification': certification,
            'budget': budget,
            'instructors': instructors,
            'survey': survey
        }
        
    except Exception as e:
        st.error(f"⚠️ 데이터 로드 중 오류가 발생했습니다.")
        with st.expander("🔍 오류 상세 정보"):
            st.error(str(e))
            
            # 엑셀 파일 시트명 확인 시도
            if 'file_path' in locals() and file_path:
                try:
                    xl_file = pd.ExcelFile(file_path)
                    st.info(f"엑셀 파일 시트 목록: {xl_file.sheet_names}")
                except Exception as sheet_error:
                    st.error(f"시트 정보 확인 실패: {str(sheet_error)}")
        
        return None

# 필터 적용 함수
def apply_filters(data):
    """필터를 적용하여 데이터를 반환"""
    filtered_data = data.copy()
    
    # session_state에서 필터 값 가져오기
    if 'filter_program' in st.session_state and st.session_state.filter_program != '전체':
        prog_id = filtered_data['program_info'][
            filtered_data['program_info']['program_name'] == st.session_state.filter_program
        ]['program_id'].values[0]
        
        # 각 데이터프레임 필터링
        filtered_data['program_info'] = filtered_data['program_info'][
            filtered_data['program_info']['program_id'] == prog_id
        ]
        filtered_data['learners'] = filtered_data['learners'][
            filtered_data['learners']['program_id'] == prog_id
        ]
        filtered_data['budget'] = filtered_data['budget'][
            filtered_data['budget']['program_id'] == prog_id
        ]
        filtered_data['instructors'] = filtered_data['instructors'][
            filtered_data['instructors']['program_id'] == prog_id
        ]
        filtered_data['survey'] = filtered_data['survey'][
            filtered_data['survey']['program_id'] == prog_id
        ]
        filtered_data['certification'] = filtered_data['certification'][
            filtered_data['certification']['program_id'] == prog_id
        ]
    
    # 회사 필터 적용
    if 'filter_companies' in st.session_state and len(st.session_state.filter_companies) > 0:
        filtered_data['learners'] = filtered_data['learners'][
            filtered_data['learners']['company'].isin(st.session_state.filter_companies)
        ]
        filtered_data['survey'] = filtered_data['survey'][
            filtered_data['survey']['company'].isin(st.session_state.filter_companies)
        ]
    
    # 기간 필터 적용
    if 'filter_months' in st.session_state and len(st.session_state.filter_months) > 0:
        filtered_data['program_info'] = filtered_data['program_info'][
            filtered_data['program_info']['program_month'].dt.strftime('%Y-%m').isin(st.session_state.filter_months)
        ]
        # 필터링된 프로그램의 ID만 유지
        valid_program_ids = filtered_data['program_info']['program_id'].unique()
        filtered_data['learners'] = filtered_data['learners'][
            filtered_data['learners']['program_id'].isin(valid_program_ids)
        ]
        filtered_data['budget'] = filtered_data['budget'][
            filtered_data['budget']['program_id'].isin(valid_program_ids)
        ]
        filtered_data['instructors'] = filtered_data['instructors'][
            filtered_data['instructors']['program_id'].isin(valid_program_ids)
        ]
        filtered_data['survey'] = filtered_data['survey'][
            filtered_data['survey']['program_id'].isin(valid_program_ids)
        ]
        filtered_data['certification'] = filtered_data['certification'][
            filtered_data['certification']['program_id'].isin(valid_program_ids)
        ]
    
    return filtered_data

# 사이드바 필터 설정
def setup_sidebar_filters(data):
    """사이드바 필터 설정"""
    st.sidebar.title("🔍 필터 옵션")
    
    # 필터 초기화 버튼
    if st.sidebar.button("🔄 필터 초기화", use_container_width=True):
        st.session_state.filter_program = '전체'
        st.session_state.filter_companies = []
        st.session_state.filter_months = []
        st.rerun()
    
    st.sidebar.markdown("---")
    
    # 프로그램 필터
    programs = ['전체'] + data['program_info']['program_name'].tolist()
    selected_program = st.sidebar.selectbox(
        "프로그램 선택", 
        programs,
        key="temp_program",
        index=programs.index(st.session_state.get('filter_program', '전체'))
    )
    
    # 회사 필터
    companies = data['learners']['company'].dropna().unique().tolist()
    selected_companies = st.sidebar.multiselect(
        "회사 선택", 
        companies,
        key="temp_companies",
        default=st.session_state.get('filter_companies', [])
    )
    
    # 기간 필터
    months = data['program_info']['program_month'].dt.strftime('%Y-%m').unique()
    selected_months = st.sidebar.multiselect(
        "월 선택", 
        months,
        key="temp_months",
        default=st.session_state.get('filter_months', [])
    )
    
    st.sidebar.markdown("---")
    
    # 적용 버튼
    if st.sidebar.button("✅ 필터 적용", type="primary", use_container_width=True):
        st.session_state.filter_program = selected_program
        st.session_state.filter_companies = selected_companies
        st.session_state.filter_months = selected_months
        st.rerun()
    
    # 현재 적용된 필터 표시
    st.sidebar.markdown("### 📌 적용된 필터")
    if 'filter_program' in st.session_state and st.session_state.filter_program != '전체':
        st.sidebar.info(f"프로그램: {st.session_state.filter_program}")
    if 'filter_companies' in st.session_state and len(st.session_state.filter_companies) > 0:
        st.sidebar.info(f"회사: {len(st.session_state.filter_companies)}개 선택")
    if 'filter_months' in st.session_state and len(st.session_state.filter_months) > 0:
        st.sidebar.info(f"기간: {len(st.session_state.filter_months)}개월 선택")
    
    return selected_program, selected_companies, selected_months

# Overview 페이지 (수정됨: 직무분야 고정, 월별 차트 수정)
def show_overview(data):
    """전체 현황 대시보드"""
    st.markdown("### 📊 전체 현황 Overview")
    
    # 데이터가 비어있는지 확인
    if len(data['program_info']) == 0:
        st.warning("선택한 필터에 해당하는 데이터가 없습니다.")
        return
    
    # KPI 계산
    total_programs = len(data['program_info'])
    total_learners = data['learners'].shape[0]
    total_budget = data['budget']['actual_budget'].sum() / 1000000 if len(data['budget']) > 0 else 0
    total_direct_cost = data['budget']['total_direct_cost'].sum() / 1000000 if len(data['budget']) > 0 else 0
    avg_satisfaction = data['survey'][data['survey']['rating'].notna()]['rating'].mean() if len(data['survey']) > 0 else 0
    
    # KPI 카드 표시
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("총 프로그램 수", f"{total_programs}개", "")
    with col2:
        st.metric("총 수강생 수", f"{total_learners}명", "")
    with col3:
        st.metric("총 예산", f"{total_budget:.1f}백만원", "")
    with col4:
        st.metric("총 직접비", f"{total_direct_cost:.1f}백만원", "")
    with col5:
        st.metric("평균 만족도", f"{avg_satisfaction:.2f}/5.0", "")
    
    st.markdown("---")
    
    # 차트 영역
    col1, col2 = st.columns(2)
    
    with col1:
        # 월별 교육 운영 현황 (수정: 1-12월 고정, 월 단위만 표시)
        # 1-12월 데이터 프레임 생성
        months_df = pd.DataFrame({'month': range(1, 13)})
        months_df['month_str'] = months_df['month'].apply(lambda x: f"{x}월")
        
        # 실제 데이터 집계
        program_monthly = data['program_info'].copy()
        program_monthly['month'] = program_monthly['program_month'].dt.month
        monthly_count = program_monthly.groupby('month').size().reset_index(name='프로그램 수')
        
        # 1-12월과 실제 데이터 병합
        months_df = months_df.merge(monthly_count, on='month', how='left')
        months_df['프로그램 수'] = months_df['프로그램 수'].fillna(0)
        
        fig1 = px.bar(months_df, x='month_str', y='프로그램 수',
                     title="월별 교육 프로그램 운영 현황",
                     color_discrete_sequence=['#ea002c'])
        fig1.update_layout(height=400, xaxis_title="",
                          xaxis={'categoryorder': 'array', 
                                 'categoryarray': ['1월','2월','3월','4월','5월','6월','7월','8월','9월','10월','11월','12월']})
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # 직무분야별 프로그램 수 (수정: 11개 직무 고정, 프로그램 수로 변경)
        job_categories = ['전략', '사업개발', '재무', 'HR', '마케팅', 'Sales', '법무', 'IP', '구매/SCM', 'SVESG', '일하는 방식']
        
        # 실제 데이터에서 직무별 프로그램 수 계산
        job_programs = data['program_info'].groupby('job_category').size().reset_index(name='프로그램 수')
        
        # 11개 카테고리 데이터프레임 생성
        job_df = pd.DataFrame({'job_category': job_categories})
        job_df = job_df.merge(job_programs, on='job_category', how='left')
        job_df['프로그램 수'] = job_df['프로그램 수'].fillna(0)
        
        fig2 = px.bar(job_df, x='job_category', y='프로그램 수',
                     title="직무분야별 프로그램 수",
                     color_discrete_sequence=['#ff5800'])
        fig2.update_layout(height=400, xaxis_title="직무분야")
        st.plotly_chart(fig2, use_container_width=True)
    
    # 프로그램 요약 테이블
    st.markdown("### 📋 프로그램 요약")
    program_summary = data['program_info'].merge(data['budget'], on='program_id')
    
    # 만족도 계산
    satisfaction_by_program = data['survey'].groupby('program_id')['rating'].mean().reset_index()
    satisfaction_by_program.columns = ['program_id', 'avg_satisfaction']
    program_summary = program_summary.merge(satisfaction_by_program, on='program_id', how='left')
    
    summary_display = program_summary[['program_name', 'job_category', 'num_learners', 
                                       'actual_budget', 'total_direct_cost', 'avg_satisfaction']].copy()
    summary_display['actual_budget'] = (summary_display['actual_budget'] / 1000000).round(1)
    summary_display['total_direct_cost'] = (summary_display['total_direct_cost'] / 1000000).round(1)
    summary_display['avg_satisfaction'] = summary_display['avg_satisfaction'].round(2)
    
    summary_display.columns = ['프로그램명', '직무분야', '수강생수', '예산(백만원)', '직접비(백만원)', '만족도']
    st.dataframe(summary_display, use_container_width=True, hide_index=True)

# 프로그램별 상세 페이지
def show_program_details(data, selected_program):
    """프로그램별 상세 분석"""
    st.markdown("### 🎓 프로그램별 상세 분석")
    
    # 필터링된 프로그램 목록만 사용
    if len(data['program_info']) == 0:
        st.warning("선택한 필터에 해당하는 프로그램이 없습니다.")
        return
    
    programs = data['program_info']['program_name'].tolist()
    
    # 필터로 특정 프로그램이 선택된 경우
    if selected_program != '전체' and selected_program in programs:
        selected_prog_name = selected_program
    else:
        # 필터링된 프로그램 목록에서 선택
        selected_prog_name = st.selectbox("분석할 프로그램 선택", programs)
    
    # 선택된 프로그램 정보 가져오기
    prog_info = data['program_info'][data['program_info']['program_name'] == selected_prog_name]
    if len(prog_info) == 0:
        st.warning("선택한 프로그램의 정보를 찾을 수 없습니다.")
        return
    
    prog_info = prog_info.iloc[0]
    prog_id = prog_info['program_id']
    
    prog_budget_df = data['budget'][data['budget']['program_id'] == prog_id]
    if len(prog_budget_df) == 0:
        st.warning("선택한 프로그램의 예산 정보를 찾을 수 없습니다.")
        return
    prog_budget = prog_budget_df.iloc[0]
    
    # 프로그램 정보 카드
    st.markdown(f"#### 📌 {selected_prog_name}")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.info(f"**담당자:** {prog_info['owner']}")
    with col2:
        st.info(f"**교육기간:** {prog_info['duration_days']}일")
    with col3:
        st.info(f"**장소:** {prog_info['venue']}")
    with col4:
        st.info(f"**대상:** {prog_info['target_company']}")
    
    st.markdown("---")
    
    # 차트 영역
    col1, col2 = st.columns(2)
    
    with col1:
        # 수강생 회사별 분포
        prog_learners = data['learners'][data['learners']['program_id'] == prog_id]
        company_dist = prog_learners['company'].value_counts().head(10)
        fig1 = px.bar(x=company_dist.values, y=company_dist.index, orientation='h',
                     title="회사별 수강생 분포",
                     labels={'x': '수강생 수', 'y': '회사'},
                     color_discrete_sequence=['#ff5800'])
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # 예산 vs 직접비 비교
        fig2 = go.Figure()
        fig2.add_trace(go.Bar(name='예산', x=['예산 vs 직접비'], 
                             y=[prog_budget['actual_budget']/1000000],
                             marker_color='#ea002c'))
        fig2.add_trace(go.Bar(name='직접비', x=['예산 vs 직접비'], 
                             y=[prog_budget['total_direct_cost']/1000000],
                             marker_color='#ffa500'))
        fig2.update_layout(title="예산 vs 직접비 비교 (백만원)",
                          barmode='group',
                          yaxis_title="금액 (백만원)")
        st.plotly_chart(fig2, use_container_width=True)
    
    # 강사 정보
    st.markdown("#### 👨‍🏫 강사진 정보")
    prog_instructors = data['instructors'][data['instructors']['program_id'] == prog_id]
    inst_display = prog_instructors[['instructor_name', 'lecture_hours', 'lecture_fee']].copy()
    inst_display['lecture_fee'] = (inst_display['lecture_fee'] / 1000000).round(1)
    inst_display.columns = ['강사명', '강의시간', '강사료(백만원)']
    st.dataframe(inst_display, use_container_width=True, hide_index=True)
    
    # 자격증 정보 (있는 경우)
    if prog_id in data['certification']['program_id'].values:
        cert_info = data['certification'][data['certification']['program_id'] == prog_id].iloc[0]
        st.markdown("#### 🏆 자격증 취득 현황")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("자격증 종류", cert_info['certification_type'])
        with col2:
            st.metric("응시자", f"{cert_info['exam_candidates']}명")
        with col3:
            pass_rate = (cert_info['exam_passed'] / cert_info['exam_candidates'] * 100)
            st.metric("합격률", f"{pass_rate:.1f}%")

# 수강생 분석 페이지
def show_learner_analysis(data):
    """수강생 분석"""
    st.markdown("### 👥 수강생 분석")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # 회사별 수강생 현황 (Top 10)
        company_counts = data['learners']['company'].value_counts().head(10)
        fig1 = px.bar(x=company_counts.values, y=company_counts.index,
                     orientation='h',
                     title="회사별 수강생 현황 (Top 10)",
                     labels={'x': '수강생 수', 'y': '회사'},
                     color=company_counts.values,
                     color_continuous_scale=['#ffa500', '#ff5800', '#ea002c'])
        fig1.update_layout(height=400)
        st.plotly_chart(fig1, use_container_width=True)
    
    with col2:
        # 직급별 분포
        level_counts = data['learners']['job_level'].value_counts()
        fig2 = px.pie(values=level_counts.values, names=level_counts.index,
                     title="직급별 분포",
                     color_discrete_sequence=['#ea002c', '#ff5800', '#ffa500'])
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)
    
    # 프로그램별 x 회사별 히트맵
    st.markdown("#### 🔥 프로그램별 x 회사별 수강생 분포")
    
    # 데이터 준비
    learner_program = data['learners'].merge(data['program_info'][['program_id', 'program_name']], on='program_id')
    heatmap_data = learner_program.groupby(['program_name', 'company']).size().reset_index(name='count')
    heatmap_pivot = heatmap_data.pivot(index='company', columns='program_name', values='count').fillna(0)
    
    fig3 = px.imshow(heatmap_pivot,
                    labels=dict(x="프로그램", y="회사", color="수강생 수"),
                    color_continuous_scale=['white', '#ffa500', '#ea002c'],
                    aspect="auto")
    fig3.update_layout(height=500)
    st.plotly_chart(fig3, use_container_width=True)
    
    # 수강생 상세 리스트
    st.markdown("#### 📋 수강생 상세 리스트")
    
    # 필터링 옵션
    col1, col2, col3 = st.columns(3)
    with col1:
        filter_company = st.selectbox("회사 필터", ['전체'] + data['learners']['company'].dropna().unique().tolist())
    with col2:
        filter_program = st.selectbox("프로그램 필터", ['전체'] + data['program_info']['program_name'].unique().tolist())
    with col3:
        filter_level = st.selectbox("직급 필터", ['전체'] + data['learners']['job_level'].dropna().unique().tolist())
    
    # 필터링 적용
    filtered_learners = data['learners'].merge(data['program_info'][['program_id', 'program_name']], on='program_id')
    
    if filter_company != '전체':
        filtered_learners = filtered_learners[filtered_learners['company'] == filter_company]
    if filter_program != '전체':
        filtered_learners = filtered_learners[filtered_learners['program_name'] == filter_program]
    if filter_level != '전체':
        filtered_learners = filtered_learners[filtered_learners['job_level'] == filter_level]
    
    display_cols = ['learner_id', 'program_name', 'company', 'dept', 'job_level']
    display_df = filtered_learners[display_cols].copy()
    display_df.columns = ['수강생ID', '프로그램', '회사', '부서', '직급']
    
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    st.info(f"총 {len(display_df)}명의 수강생이 검색되었습니다.")

# 예산 분석 페이지 (수정됨: KPI 카드 수정)
def show_budget_analysis(data):
    """예산 분석"""
    st.markdown("### 💰 예산 분석")
    
    # 상단 요약 카드 (수정: 개발비, 강사료, 예비비, 평균예산 추가)
    total_budget = data['budget']['actual_budget'].sum() / 1000000
    total_direct = data['budget']['total_direct_cost'].sum() / 1000000
    total_dev_cost = data['budget']['dev_cost'].sum() / 1000000
    total_instructor_fee = data['budget']['instructor_fee'].sum() / 1000000
    total_reserve_fund = data['budget']['reserve_fund'].sum() / 1000000
    avg_budget = total_budget / len(data['program_info'])
    
    # 첫 번째 줄: 주요 지표
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("총 예산", f"{total_budget:.1f}백만원", "")
    with col2:
        st.metric("총 직접비", f"{total_direct:.1f}백만원", "")
    with col3:
        st.metric("평균 예산", f"{avg_budget:.1f}백만원", "프로그램당")
    with col4:
        budget_ratio = (total_budget/(total_budget+total_direct)*100)
        st.metric("예산 비율", f"{budget_ratio:.1f}%", "")
    
    # 두 번째 줄: 예산 세부 항목
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("총 개발비", f"{total_dev_cost:.1f}백만원", 
                 f"{(total_dev_cost/total_budget*100):.1f}%")
    with col2:
        st.metric("총 강사료", f"{total_instructor_fee:.1f}백만원",
                 f"{(total_instructor_fee/total_budget*100):.1f}%")
    with col3:
        st.metric("총 예비비", f"{total_reserve_fund:.1f}백만원",
                 f"{(total_reserve_fund/total_budget*100):.1f}%")
    
    st.markdown("---")
    
    # 프로그램별 예산 vs 직접비 비교
    st.markdown("#### 📊 프로그램별 예산 vs 직접비 비교")
    
    # budget_comparison 생성 시 program_id 유지
    budget_comparison = data['budget'].copy()
    budget_comparison = budget_comparison.merge(data['program_info'][['program_id', 'program_name']], on='program_id', how='left')
    
    fig1 = go.Figure()
    fig1.add_trace(go.Bar(name='예산', x=budget_comparison['program_name'], 
                         y=budget_comparison['actual_budget']/1000000,
                         marker_color='#ea002c',
                         text=[f"{x:.1f}M" for x in budget_comparison['actual_budget']/1000000],
                         textposition='auto'))
    fig1.add_trace(go.Bar(name='직접비', x=budget_comparison['program_name'], 
                         y=budget_comparison['total_direct_cost']/1000000,
                         marker_color='#ffa500',
                         text=[f"{x:.1f}M" for x in budget_comparison['total_direct_cost']/1000000],
                         textposition='auto'))
    fig1.update_layout(title="프로그램별 예산 vs 직접비 (백만원)",
                      barmode='group',
                      yaxis_title="금액 (백만원)",
                      height=400)
    st.plotly_chart(fig1, use_container_width=True)
    
    # 예산 구성 분석
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("#### 💼 예산 항목별 분포")
        
        fig2 = px.pie(values=[total_dev_cost, total_instructor_fee, total_reserve_fund],
                     names=['개발비', '강사료', '예비비'],
                     color_discrete_sequence=['#ea002c', '#ff5800', '#ffa500'],
                     hole=0.4)
        fig2.update_traces(textposition='inside', textinfo='percent+label')
        fig2.update_layout(height=400)
        st.plotly_chart(fig2, use_container_width=True)
    
    with col2:
        st.markdown("#### 💵 직접비 효율성 분석")
        
        # 1인당 직접비
        per_person_cost = data['budget']['direct_cost'].mean() / 1000
        st.metric("1인당 직접비", f"{per_person_cost:.0f}천원", "")
        
        # 프로그램별 직접비 효율성 테이블 - 처음부터 다시 생성
        efficiency_df = data['budget'].copy()
        efficiency_df = efficiency_df.merge(data['program_info'][['program_id', 'program_name', 'num_learners']], 
                                          on='program_id')
        efficiency_df['직접비_비율'] = (efficiency_df['total_direct_cost'] / 
                                    (efficiency_df['actual_budget'] + efficiency_df['total_direct_cost']) * 100)
        
        display_efficiency = efficiency_df[['program_name', 'num_learners', 'direct_cost', 
                                           'total_direct_cost', '직접비_비율']].copy()
        display_efficiency['direct_cost'] = (display_efficiency['direct_cost'] / 1000).round(0)
        display_efficiency['total_direct_cost'] = (display_efficiency['total_direct_cost'] / 1000000).round(1)
        display_efficiency['직접비_비율'] = display_efficiency['직접비_비율'].round(1)
        display_efficiency.columns = ['프로그램', '수강생수', '1인당(천원)', '총액(백만원)', '비율(%)']
        
        st.dataframe(display_efficiency, use_container_width=True, hide_index=True)
    
    # 전체 비용 구조
    st.markdown("#### 📈 예산 항목별 분포 상세")
    
    # Stacked bar chart
    fig3 = go.Figure()
    
    programs = budget_comparison['program_name'].unique()
    dev_cost_values = []
    instructor_fee_values = []
    reserve_fund_values = []
    
    for prog in programs:
        prog_data = budget_comparison[budget_comparison['program_name'] == prog]
        dev_cost_values.append(prog_data['dev_cost'].values[0] / 1000000)
        instructor_fee_values.append(prog_data['instructor_fee'].values[0] / 1000000)
        reserve_fund_values.append(prog_data['reserve_fund'].values[0] / 1000000)
    
    fig3.add_trace(go.Bar(name='개발비', x=programs, y=dev_cost_values,
                         marker_color='#ea002c'))
    fig3.add_trace(go.Bar(name='강사료', x=programs, y=instructor_fee_values,
                         marker_color='#ff5800'))
    fig3.add_trace(go.Bar(name='예비비', x=programs, y=reserve_fund_values,
                         marker_color='#ffa500'))
    
    fig3.update_layout(barmode='stack',
                      title='프로그램별 예산 항목별 분포 상세 (백만원)',
                      yaxis_title='금액 (백만원)',
                      height=400)
    st.plotly_chart(fig3, use_container_width=True)
    
    # 강사료 상세 분석
    st.markdown("#### 👨‍🏫 강사료 분석")
    instructor_analysis = data['instructors'].merge(
        data['program_info'][['program_id', 'program_name']], on='program_id')
    
    # 프로그램별 강사료 총액
    prog_instructor_fee = instructor_analysis.groupby('program_name')['lecture_fee'].sum().reset_index()
    prog_instructor_fee['lecture_fee'] = (prog_instructor_fee['lecture_fee'] / 1000000).round(1)
    
    col1, col2 = st.columns(2)
    with col1:
        fig4 = px.bar(prog_instructor_fee, x='program_name', y='lecture_fee',
                     title="프로그램별 강사료 총액 (백만원)",
                     color_discrete_sequence=['#ff5800'])
        st.plotly_chart(fig4, use_container_width=True)
    
    with col2:
        # 시간당 단가 분석
        instructor_analysis['hourly_rate'] = instructor_analysis['lecture_fee'] / instructor_analysis['lecture_hours'] / 10000
        avg_hourly = instructor_analysis.groupby('program_name')['hourly_rate'].mean().reset_index()
        fig5 = px.bar(avg_hourly, x='program_name', y='hourly_rate',
                     title="프로그램별 평균 시간당 강사료 (만원)",
                     color_discrete_sequence=['#ffa500'])
        st.plotly_chart(fig5, use_container_width=True)

# 만족도 분석 페이지 (수정됨: 프로그램별 선택 기능 추가)
def show_satisfaction_analysis(data):
    """만족도 분석"""
    st.markdown("### ⭐ 만족도 분석")
    
    # 프로그램 선택 기능
    programs = ['전체'] + data['program_info']['program_name'].tolist()
    selected_prog_for_satisfaction = st.selectbox(
        "분석할 프로그램 선택", 
        programs, 
        key="satisfaction_program_select"
    )
    
    # 데이터 필터링
    if selected_prog_for_satisfaction == '전체':
        satisfaction_data = data['survey'][data['survey']['rating'].notna()]
        all_survey_data = data['survey']
        prog_label = "전체"
    else:
        prog_id = data['program_info'][data['program_info']['program_name'] == selected_prog_for_satisfaction]['program_id'].values[0]
        satisfaction_data = data['survey'][(data['survey']['program_id'] == prog_id) & (data['survey']['rating'].notna())]
        all_survey_data = data['survey'][data['survey']['program_id'] == prog_id]
        prog_label = selected_prog_for_satisfaction
    
    # 전체 만족도 계산
    overall_satisfaction = satisfaction_data['rating'].mean()
    
    # 큰 카드로 만족도 표시
    st.markdown(
        f"""
        <div style='text-align: center; padding: 30px; background: linear-gradient(135deg, #ea002c, #ff5800); 
                    border-radius: 20px; margin: 20px 0;'>
            <h1 style='color: white; margin: 0;'>{prog_label} 만족도</h1>
            <h1 style='color: white; font-size: 60px; margin: 10px 0;'>{overall_satisfaction:.2f} / 5.0</h1>
            <p style='color: white; font-size: 20px;'>{'⭐' * int(overall_satisfaction)}</p>
        </div>
        """, unsafe_allow_html=True
    )
    
    if selected_prog_for_satisfaction == '전체':
        # 전체 선택시: 프로그램별 비교, 회사별 만족도 표시
        col1, col2 = st.columns(2)
        
        with col1:
            # 프로그램별 평균 만족도
            prog_satisfaction = data['survey'][data['survey']['rating'].notna()].merge(
                data['program_info'][['program_id', 'program_name']], 
                on='program_id')
            prog_avg = prog_satisfaction.groupby('program_name')['rating'].mean().reset_index()
            
            fig1 = px.bar(prog_avg, x='rating', y='program_name', orientation='h',
                         title="프로그램별 만족도 비교",
                         color='rating',
                         color_continuous_scale=['#ffa500', '#ff5800', '#ea002c'],
                         range_x=[0, 5])
            fig1.update_layout(height=400)
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # 질문별 평균 점수
            question_avg = satisfaction_data.groupby('question_id').agg({
                'rating': 'mean',
                'question_text': 'first'
            }).reset_index()
            
            question_avg['question_short'] = question_avg['question_id'].map({
                'Q1': '전반적 만족도',
                'Q2': '추천 의향',
                'Q3': '실무 도움도'
            })
            
            fig2 = go.Figure(go.Scatterpolar(
                r=question_avg['rating'],
                theta=question_avg['question_short'],
                fill='toself',
                marker_color='#ea002c',
                name='만족도'
            ))
            fig2.update_layout(
                polar=dict(
                    radialaxis=dict(
                        visible=True,
                        range=[0, 5]
                    )),
                showlegend=False,
                title="질문별 만족도 레이더 차트"
            )
            st.plotly_chart(fig2, use_container_width=True)
        
        # 회사별 만족도 분포
        st.markdown("#### 🏢 회사별 만족도 분포")
        
        company_satisfaction = satisfaction_data.groupby('company')['rating'].agg(['mean', 'std', 'count']).reset_index()
        company_satisfaction = company_satisfaction[company_satisfaction['count'] >= 5]  # 5개 이상 응답만
        company_satisfaction = company_satisfaction.sort_values('mean', ascending=False).head(15)
        
        fig3 = go.Figure()
        fig3.add_trace(go.Bar(
            x=company_satisfaction['company'],
            y=company_satisfaction['mean'],
            error_y=dict(type='data', array=company_satisfaction['std']),
            marker_color='#ff5800',
            name='평균 만족도'
        ))
        fig3.update_layout(
            title="회사별 만족도 (상위 15개사)",
            yaxis_title="만족도",
            xaxis_tickangle=-45,
            height=500
        )
        st.plotly_chart(fig3, use_container_width=True)
        
    else:
        # 개별 프로그램 선택시: 상세 분석
        
        # 객관식 문항별 상세 평균
        st.markdown("#### 📊 객관식 문항별 평균 평점")
        
        objective_questions = satisfaction_data[satisfaction_data['question_type'] == '객관식']
        question_details = objective_questions.groupby(['question_id', 'question_text']).agg({
            'rating': ['mean', 'std', 'count']
        }).reset_index()
        question_details.columns = ['question_id', 'question_text', 'mean', 'std', 'count']
        
        # 문항별 상세 카드
        for _, row in question_details.iterrows():
            col1, col2, col3, col4 = st.columns([3, 1, 1, 1])
            with col1:
                # 질문 텍스트 간소화
                question_display = row['question_text'].split(']')[1].strip() if ']' in row['question_text'] else row['question_text']
                st.write(f"**{row['question_id']}**: {question_display}")
            with col2:
                st.metric("평균", f"{row['mean']:.2f}")
            with col3:
                st.metric("표준편차", f"{row['std']:.2f}")
            with col4:
                st.metric("응답수", f"{int(row['count'])}명")
        
        # 객관식 점수 분포 시각화
        col1, col2 = st.columns(2)
        
        with col1:
            # 질문별 평균 점수 비교
            question_scores = question_details.copy()
            question_scores['question_short'] = question_scores['question_id'].map({
                'Q1': '전반적 만족도',
                'Q2': '추천 의향',
                'Q3': '실무 도움도'
            })
            
            fig1 = px.bar(question_scores, x='question_short', y='mean',
                         title="객관식 문항별 평균 점수",
                         color='mean',
                         color_continuous_scale=['#ffa500', '#ff5800', '#ea002c'],
                         range_y=[0, 5],
                         text='mean')
            fig1.update_traces(texttemplate='%{text:.2f}', textposition='outside')
            fig1.update_layout(height=400, xaxis_title="", yaxis_title="평균 점수")
            st.plotly_chart(fig1, use_container_width=True)
        
        with col2:
            # 만족도 점수 분포
            fig2 = px.histogram(satisfaction_data, x='rating',
                              title="만족도 점수 분포",
                              nbins=5,
                              color_discrete_sequence=['#ea002c'])
            fig2.update_layout(height=400, 
                             xaxis_title="만족도 점수",
                             yaxis_title="응답 수",
                             xaxis=dict(tickmode='linear', tick0=1, dtick=1))
            st.plotly_chart(fig2, use_container_width=True)
        
        # 주관식 응답 분석
        st.markdown("#### 💬 주관식 문항 의견 요약")
        
        subjective_data = all_survey_data[(all_survey_data['question_type'] == '주관식') & 
                                          (all_survey_data['comment'].notna())]
        
        if len(subjective_data) > 0:
            # 주관식 질문별로 그룹화
            subjective_questions = subjective_data['question_text'].unique()
            
            for question in subjective_questions:
                question_comments = subjective_data[subjective_data['question_text'] == question]['comment']
                
                if len(question_comments) > 0:
                    # 질문 표시
                    question_display = question.split(']')[1].strip() if ']' in question else question
                    st.markdown(f"**📝 {question_display}**")
                    
                    # 키워드 분석 (간단한 방식)
                    # 모든 코멘트를 하나로 합치고 단어 추출
                    all_comments = ' '.join(question_comments.values)
                    
                    # 의미있는 키워드 추출 (2글자 이상)
                    words = re.findall(r'[가-힣]+', all_comments)
                    words = [word for word in words if len(word) >= 2]
                    
                    # 빈도수 계산
                    word_freq = Counter(words)
                    top_keywords = word_freq.most_common(10)
                    
                    # 자주 언급되는 키워드 표시
                    if top_keywords:
                        keyword_html = ""
                        for word, freq in top_keywords[:5]:
                            if freq >= 3:  # 3회 이상 언급된 키워드만
                                size = min(30, 15 + freq * 2)  # 빈도에 따라 크기 조정
                                keyword_html += f'<span style="font-size: {size}px; color: #ea002c; margin: 5px; font-weight: bold;">{word}</span> '
                        
                        if keyword_html:
                            st.markdown(f"**자주 언급된 키워드:** {keyword_html}", unsafe_allow_html=True)
                    
                    # 대표 의견 표시 (샘플)
                    st.markdown("**주요 의견:**")
                    sample_size = min(5, len(question_comments))
                    for comment in question_comments.sample(sample_size).values:
                        # 자주 언급된 키워드 강조
                        highlighted_comment = comment
                        for word, freq in top_keywords[:5]:
                            if freq >= 3:
                                highlighted_comment = highlighted_comment.replace(
                                    word, 
                                    f"**<span style='color: #ea002c;'>{word}</span>**"
                                )
                        st.markdown(f"• {highlighted_comment}", unsafe_allow_html=True)
                    
                    st.markdown("")  # 구분을 위한 빈 줄
        
        else:
            st.info("주관식 응답이 없습니다.")
        
        # 회사별 만족도 (해당 프로그램만)
        st.markdown(f"#### 🏢 {selected_prog_for_satisfaction} - 회사별 만족도")
        
        company_prog_satisfaction = satisfaction_data.groupby('company')['rating'].agg(['mean', 'count']).reset_index()
        company_prog_satisfaction = company_prog_satisfaction[company_prog_satisfaction['count'] >= 3]  # 3개 이상 응답만
        company_prog_satisfaction = company_prog_satisfaction.sort_values('mean', ascending=False)
        
        if len(company_prog_satisfaction) > 0:
            fig3 = px.bar(company_prog_satisfaction, x='company', y='mean',
                         title=f"회사별 만족도 평균",
                         color='mean',
                         color_continuous_scale=['#ffa500', '#ff5800', '#ea002c'],
                         range_y=[0, 5])
            fig3.update_layout(xaxis_tickangle=-45, height=400,
                             xaxis_title="회사", yaxis_title="평균 만족도")
            st.plotly_chart(fig3, use_container_width=True)
    
    # 전체 주관식 응답 요약 (전체 선택시에만)
    if selected_prog_for_satisfaction == '전체':
        comments = all_survey_data[all_survey_data['comment'].notna()]['comment']
        if len(comments) > 0:
            st.markdown("#### 💬 전체 프로그램 주요 피드백")
            
            # 키워드 분석
            all_comments = ' '.join(comments.values)
            words = re.findall(r'[가-힣]+', all_comments)
            words = [word for word in words if len(word) >= 2]
            word_freq = Counter(words)
            top_keywords = word_freq.most_common(15)
            
            # 워드 클라우드 스타일로 키워드 표시
            keyword_html = "<div style='text-align: center; padding: 20px; background-color: #f9f9f9; border-radius: 10px;'>"
            for word, freq in top_keywords:
                if freq >= 5:
                    size = min(35, 12 + freq)
                    color = '#ea002c' if freq >= 10 else '#ff5800' if freq >= 7 else '#ffa500'
                    keyword_html += f'<span style="font-size: {size}px; color: {color}; margin: 8px; display: inline-block; font-weight: bold;">{word}</span> '
            keyword_html += "</div>"
            
            st.markdown("**자주 언급된 표현:**", unsafe_allow_html=True)
            st.markdown(keyword_html, unsafe_allow_html=True)
            
            # 샘플 코멘트 표시
            st.info("📝 수강생 주요 의견 (샘플)")
            sample_comments = comments.sample(min(5, len(comments)))
            for comment in sample_comments:
                st.write(f"• {comment}")

# 메인 함수
def main():
    # 세션 상태 초기화
    if 'filter_program' not in st.session_state:
        st.session_state.filter_program = '전체'
    if 'filter_companies' not in st.session_state:
        st.session_state.filter_companies = []
    if 'filter_months' not in st.session_state:
        st.session_state.filter_months = []
    
    # 데이터 로드
    data = load_data()
    
    if data is None:
        # 파일을 찾을 수 없을 때 바로 업로드 인터페이스 제공 (오류 메시지 없이)
        st.title("📚 2025년 성장지원 워크샵 교육과정 대시보드")
        st.markdown("### 📁 데이터 파일 업로드")
        st.info("교육 데이터가 포함된 엑셀 파일을 업로드해주세요.")
        
        uploaded_file = st.file_uploader("Excel 파일 선택 (.xlsx)", type=['xlsx'])
        
        if uploaded_file is not None:
            try:
                # 업로드된 파일로부터 데이터 로드
                data = {
                    'program_info': pd.read_excel(uploaded_file, sheet_name='Program_Info'),
                    'learners': pd.read_excel(uploaded_file, sheet_name='Learners'),
                    'certification': pd.read_excel(uploaded_file, sheet_name='Certification'),
                    'budget': pd.read_excel(uploaded_file, sheet_name='Budget'),
                    'instructors': pd.read_excel(uploaded_file, sheet_name='Instructors'),
                    'survey': pd.read_excel(uploaded_file, sheet_name='Survey')
                }
                
                # 날짜 형식 변환
                data['program_info']['program_month'] = pd.to_datetime(data['program_info']['program_month'])
                
                # 예산 계산 추가
                data['budget']['actual_budget'] = data['budget']['dev_cost'] + data['budget']['instructor_fee'] + data['budget']['reserve_fund']
                
                # 직접비 총액 계산
                for idx, row in data['budget'].iterrows():
                    prog = data['program_info'][data['program_info']['program_id'] == row['program_id']].iloc[0]
                    data['budget'].loc[idx, 'total_direct_cost'] = row['direct_cost'] * prog['num_learners']
                
                st.success("✅ 파일이 성공적으로 로드되었습니다!")
                st.balloons()
            except Exception as e:
                st.error(f"⚠️ 파일 형식이 올바르지 않습니다. 확인 후 다시 시도해주세요.")
                st.caption(f"오류 상세: {str(e)}")
                return
        else:
            return
    
    # 타이틀
    st.title("📚 2025년 성장지원 워크샵 교육과정 대시보드")
    
    # 사이드바 필터 설정 (원본 데이터 사용)
    selected_program, selected_companies, selected_months = setup_sidebar_filters(data)
    
    # 필터가 적용된 데이터 가져오기
    filtered_data = apply_filters(data)
    
    # 필터 적용 알림
    if (st.session_state.get('filter_program', '전체') != '전체' or 
        len(st.session_state.get('filter_companies', [])) > 0 or 
        len(st.session_state.get('filter_months', [])) > 0):
        st.info("🔍 필터가 적용되었습니다. 좌측 사이드바에서 필터를 변경할 수 있습니다.")
    
    # 탭 생성 (페이지 선택을 탭으로 변경)
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🏠 Overview", 
        "🎓 프로그램별 상세", 
        "👥 수강생 분석", 
        "💰 예산 분석", 
        "⭐ 만족도 분석"
    ])
    
    # 각 탭에 해당 페이지 내용 표시 (필터링된 데이터 사용)
    with tab1:
        if len(filtered_data['program_info']) > 0:
            show_overview(filtered_data)
        else:
            st.warning("⚠️ 선택한 필터 조건에 해당하는 데이터가 없습니다.")
    
    with tab2:
        if len(filtered_data['program_info']) > 0:
            show_program_details(filtered_data, st.session_state.get('filter_program', '전체'))
        else:
            st.warning("⚠️ 선택한 필터 조건에 해당하는 프로그램이 없습니다.")
    
    with tab3:
        if len(filtered_data['learners']) > 0:
            show_learner_analysis(filtered_data)
        else:
            st.warning("⚠️ 선택한 필터 조건에 해당하는 수강생이 없습니다.")
    
    with tab4:
        if len(filtered_data['budget']) > 0:
            show_budget_analysis(filtered_data)
        else:
            st.warning("⚠️ 선택한 필터 조건에 해당하는 예산 정보가 없습니다.")
    
    with tab5:
        if len(filtered_data['survey']) > 0:
            show_satisfaction_analysis(filtered_data)
        else:
            st.warning("⚠️ 선택한 필터 조건에 해당하는 만족도 데이터가 없습니다.")
    
    # 푸터
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #888; padding: 20px;'>
            <p>© 2025 mySUNI 성장지원 워크샵 대시보드 | mySUNI 성장지원</p>
            <p>문의: suyoung.park@sk.com | 내선: 000-000-0000</p>
        </div>
        """, unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()





