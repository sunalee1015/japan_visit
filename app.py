import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 기본 설정
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
st.set_page_config(page_title="일본인 관광 지방분산 대시보드", page_icon="🗾", layout="wide")

# 색상 팔레트
SEOUL = "#C00000"
REGION = "#2E75B6"
PRIMARY = "#1F3864"
LIGHT = "#D6E4F0"

# 파일 경로 (로컬 및 Streamlit Cloud 환경 완벽 대응 - 전방위 탐색 로직)
def find_data_file():
    # 탐색 후보 루트 디렉토리들
    search_roots = [
        os.getcwd(),                                     # 현재 작업 디렉토리
        os.path.dirname(os.path.abspath(__file__)),      # 스크립트 위치 (src/)
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))), # 프로젝트 루트
        "/mount/src"                                     # Streamlit Cloud 고정 루트
    ]
    
    # 1단계: 각 루트에서 직접적으로 data/ 하위를 먼저 탐색
    for root in search_roots:
        target_dir = os.path.join(root, "data")
        if not os.path.exists(target_dir):
            # 혹시 서브폴더(예: japan_visit/data) 내에 있을 경우를 위해 현재 폴더명을 포함해 탐색
            for sub in [d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))]:
                if os.path.exists(os.path.join(root, sub, "data")):
                    target_dir = os.path.join(root, sub, "data")
                    break
        
        if os.path.exists(target_dir):
            files = [f for f in os.listdir(target_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
            if files:
                return os.path.join(target_dir, files[0])

    # 2단계: 최후의 수단 - 전체 하위 디렉토리에서 'data' 폴더 강제 탐색
    for root in search_roots:
        if os.path.exists(root):
            for r, dirs, files in os.walk(root):
                if "data" in dirs:
                    d_path = os.path.join(r, "data")
                    f_list = [f for f in os.listdir(d_path) if f.endswith('.xlsx') and not f.startswith('~$')]
                    if f_list:
                        return os.path.join(d_path, f_list[0])

    # 3단계: 기본값 반환 (여전히 못 찾은 경우)
    return os.path.join(os.getcwd(), "data", "Total_Dataset_analysis_20260321.xlsx")

DATA_PATH = find_data_file()

@st.cache_data
def load_data():
    if not os.path.exists(DATA_PATH):
        st.error(f"❌ 데이터 파일을 찾을 수 없습니다.")
        st.info(f"**현재 탐색된 경로:** {DATA_PATH}")
        st.info(f"**현재 작업 디렉토리(CWD):** {os.getcwd()}")
        st.info(f"**실행 파일 위치:** {os.path.abspath(__file__)}")
        return None
    
    xls = pd.ExcelFile(DATA_PATH)
    data = {}

    def get_df(sheet_name, iloc_slice=None):
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        if iloc_slice:
            df = df.iloc[iloc_slice]
        df = df.reset_index(drop=True)
        
        # 중복/NaN 컬럼 처리 (Plotly/Narwhals 대응)
        new_cols = []
        for i, col in enumerate(df.iloc[0]):
            if pd.isna(col) or str(col).strip() == "":
                new_cols.append(f"Unnamed_{i}")
            else:
                new_cols.append(str(col))
        
        df.columns = new_cols
        df = df.iloc[1:].reset_index(drop=True)
        return df

    # 시트 매핑 업데이트 (신규 P-시리즈 명칭 반영)
    data['s1'] = get_df("P1-1 교통수단별 입국자수", slice(2, None))
    data['s2'] = get_df("P1-2 성별·연령·목적별 통계", slice(2, None))
    data['s5'] = get_df("P2-3 광역별 연도별 방문수", slice(2, None))
    data['s6'] = get_df("P2-4 광역별 월별 방문수", slice(2, None))
    data['s9'] = get_df("P3-1 신용카드 소비 월별추이", slice(2, None))
    data['s11'] = get_df("P3-3 일본 신용카드 업종별", slice(2, 9))
    
    # 시트 P3-4: 지역별 소비비율 통합
    df12_raw = pd.read_excel(xls, sheet_name="P3-4 지역별 소비비율 통합", header=None)
    data['s12_global_card'] = df12_raw.iloc[3:21].reset_index(drop=True)
    data['s12_jp_card'] = df12_raw.iloc[24:42].reset_index(drop=True)
    
    def fix_cols_12(df):
        df.columns = ['시도명', '2024년', '2025년', '2026년'][:len(df.columns)]
        return df
    data['s12_global_card'] = fix_cols_12(data['s12_global_card'])
    data['s12_jp_card'] = fix_cols_12(data['s12_jp_card'])
    
    # 시트 P3-5: 서울 구별 소비비율 비교
    df13_raw = pd.read_excel(xls, sheet_name="P3-5 서울 구별 소비비율 비교", header=None)
    data['s13_global'] = df13_raw.iloc[3:28].reset_index(drop=True)
    data['s13_jp'] = df13_raw.iloc[31:56].reset_index(drop=True)
    for k in ['s13_global', 's13_jp']:
        cols = ['자치구', '2024년', '2025년', '2026년', '평균(%)']
        data[k].columns = cols[:len(data[k].columns)]

    data['s14'] = get_df("P4-1 유튜브 키워드 요약", slice(2, 10))
    data['s16'] = get_df("P4-3 유튜브 월별 업로드 트렌드", slice(2, None))
    data['s18'] = get_df("P4-5 관광 SNS 키워드 월별", slice(2, 12))
    
    # 시트 P5-2: 세계한류감도
    df20_raw = pd.read_excel(xls, sheet_name="P5-2 세계한류감도(일본중심)", header=None)
    data['s20_trend'] = df20_raw.iloc[3:12].reset_index(drop=True)
    data['s20_radar'] = df20_raw.iloc[15:23].reset_index(drop=True)

    data['s21'] = get_df("P4-6 K-Wave 주간키워드", slice(2, 33))
    return data

data_dict = load_data()

# 사이드바 필터
st.sidebar.header("🗾 필터 설정")
years = st.sidebar.multiselect("연도 선택", ["2023", "2024", "2025", "2026"], default=["2024", "2025"])
region_filt = st.sidebar.radio("지역 구분", ["전국", "서울만", "지방만"])
gender_filt = st.sidebar.radio("성별", ["전체", "여성", "남성"])
age_filt = st.sidebar.multiselect("연령대", ["20대", "30대", "40대", "50대+"], default=["20대", "30대", "40대"])

st.sidebar.markdown("---")
st.sidebar.caption("출처: 한국관광데이터랩, 법무부, YouTube Data API")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 KPI 요약", "🗺️ 지역 분석", "💳 소비 패턴", "📱 콘텐츠 트렌드", "🎯 전략 인사이트"])

if data_dict:
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TAB 1: KPI 요약
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with tab1:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("외국인 소비 YoY", "+21.1%", delta="+21.1%")
        c2.metric("서울 간편결제 집중", "90.5%", delta="-2.4%", delta_color="inverse")
        c3.metric("일본인 서울 방문비율", "44.1%")
        c4.metric("여성 20~30대 비중", "31.9%", delta="+관광목적")
        c5.metric("K-공연관람 SNS", "155,400건")

        col1, col2 = st.columns([6, 4])
        with col1:
            st.subheader("🗓️ 월별 신용카드 소비 추이")
            df9 = data_dict['s9'].copy()
            df9 = df9[df9['연도'].astype(str).str.contains('|'.join(years))]
            fig1 = px.line(df9, x="월", y="조회기간 소비액 (원)", color="연도", template="plotly_white")
            st.plotly_chart(fig1, use_container_width=True)
            with st.expander("📋 원본 데이터 보기"): st.dataframe(df9)

        with col2:
            st.subheader("💡 소비 지표 요약")
            st.metric("최고월 소비", "1.2조 원", "2024.03")
            st.metric("평균 소비", "0.95조 원")
        
        col3, col4 = st.columns(2)
        with col3:
            st.subheader("🔥 성별×연령 히트맵")
            df2 = data_dict['s2'].copy()
            hm_data = df2.groupby(['성별', '연령별'])['인원수'].sum().unstack()
            fig2 = px.imshow(hm_data, color_continuous_scale="Blues", template="plotly_white")
            st.plotly_chart(fig2, use_container_width=True)
            with st.expander("📋 원본 데이터 보기"): st.dataframe(df2)
        
        with col4:
            st.subheader("✈️ 교통수단별 입국자")
            df1 = data_dict['s1'].copy()
            df1_m = df1.melt(id_vars=['기준연월'], var_name='교통수단', value_name='명')
            fig3 = px.bar(df1_m, x="기준연월", y="명", color="교통수단", barmode="stack", template="plotly_white")
            st.plotly_chart(fig3, use_container_width=True)
            with st.expander("📋 원본 데이터 보기"): st.dataframe(df1)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TAB 2: 지역 분석
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with tab2:
        col1, col2 = st.columns([7, 3])
        with col1:
            st.subheader("📍 광역별 방문객 현황")
            df5 = data_dict['s5'].copy()
            df5['color'] = df5.iloc[:, 0].apply(lambda x: SEOUL if "서울" in str(x) else REGION)
            fig4 = px.bar(df5, y=df5.columns[0], x=df5.columns[-1], orientation='h', color='color', color_discrete_map="identity", template="plotly_white")
            st.plotly_chart(fig4, use_container_width=True)
            with st.expander("📋 원본 데이터 보기"): st.dataframe(df5)
        
        with col2:
            st.subheader("🎯 서울 집중도")
            fig_g = go.Figure(go.Indicator(mode="gauge+number", value=35.8, gauge={'axis': {'range': [0, 100]}, 'threshold': {'line': {'color': "red", 'width': 4}, 'value': 35.8}}))
            st.plotly_chart(fig_g, use_container_width=True)

        st.subheader("📅 광역별 월별 방문객 밀도")
        df6 = data_dict['s6'].set_index(data_dict['s6'].columns[0])
        fig5 = px.imshow(df6, color_continuous_scale="RdBu_r", template="plotly_white")
        st.plotly_chart(fig5, use_container_width=True)
        with st.expander("📋 원본 데이터 보기"): st.dataframe(df6)

        st.subheader("💳 글로벌 vs 일본 소비비교")
        df12_m = pd.merge(data_dict['s12_global_card'][['시도명', '2025년']], data_dict['s12_jp_card'][['시도명', '2025년']], on='시도명', suffixes=('_글로벌', '_일본'))
        fig6 = px.bar(df12_m, x='시도명', y=['2025년_글로벌', '2025년_일본'], barmode='group', template="plotly_white")
        st.plotly_chart(fig6, use_container_width=True)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TAB 3: 소비 패턴
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with tab3:
        col1, col2 = st.columns([4, 6])
        with col1:
            st.subheader("🍕 업종별 소비 (2025년)")
            df11 = data_dict['s11'].copy()
            # 데이터 정제 (% 제거 및 숫자 변환)
            for col in df11.columns[1:4]:
                df11[col] = pd.to_numeric(df11[col].astype(str).str.replace(r'\(%\)', '', regex=True), errors='coerce')
            
            fig7 = px.pie(df11, values=df11.columns[2], names=df11.columns[0], hole=0.5, template="plotly_white")
            st.plotly_chart(fig7, use_container_width=True)
            st.warning("⚠️ 의료웰니스 16.2% — 온천·한방·스파 등 지방 분산 핵심 업종")

        with col2:
            st.subheader("📈 연도별 업종 추이(%)")
            fig8 = px.bar(df11, x=df11.columns[0], y=df11.columns[1:4], barmode="group", text_auto='.1f', template="plotly_white")
            st.plotly_chart(fig8, use_container_width=True)
        
        st.subheader("🏘️ 서울 구별 소비 비중 (글로벌 vs 일본)")
        df13 = pd.merge(data_dict['s13_global'], data_dict['s13_jp'], on='자치구', suffixes=('_글로벌', '_일본')).iloc[:15]
        # 숫자 변환 보장
        val_cols = ['평균(%)_글로벌', '평균(%)_일본']
        for c in val_cols: df13[c] = pd.to_numeric(df13[c], errors='coerce')
        
        fig9 = px.bar(df13, y='자치구', x=val_cols, barmode='group', orientation='h', text_auto='.1f', template="plotly_white")
        fig9.add_annotation(y="중구", x=df13['평균(%)_일본'].max(), text="명동 집중", showarrow=True)
        st.plotly_chart(fig9, use_container_width=True)
        with st.expander("📋 원본 데이터 보기"): st.dataframe(df13)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TAB 4: 콘텐츠 트렌드
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with tab4:
        st.subheader("📱 SNS 키워드 누적 추이")
        df18 = data_dict['s18'].copy()
        df18_melt = df18.melt(id_vars=[df18.columns[1]], value_vars=df18.columns[2:], var_name='월', value_name='건')
        fig10 = px.area(df18_melt, x='월', y='건', color=df18.columns[1], template="plotly_white")
        st.plotly_chart(fig10, use_container_width=True)
        with st.expander("📋 원본 데이터 보기"): st.dataframe(df18)

        st.subheader("🔮 K-Wave 트렌드 및 유튜브")
        c1, c2 = st.columns(2)
        with c1:
            df14 = data_dict['s14'].copy()
            fig14 = px.bar(df14, x=df14.columns[1], y=df14.columns[0], orientation='h', title="유튜브 키워드 조회수")
            st.plotly_chart(fig14, use_container_width=True)
        with c2:
            df16 = data_dict['s16'].copy()
            fig16 = px.line(df16, x=df16.columns[0], y=df16.columns[1], markers=True, title="유튜브 업로드 추이")
            st.plotly_chart(fig16, use_container_width=True)

    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # TAB 5: 전략 인사이트
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    with tab5:
        st.header("🎯 일본인 관광 활성화를 위한 전략 제안")
        
        # 전략 카드 섹션 (기반 코드 강화)
        st.subheader("💡 분야별 집중 전략")
        st.info("데이터 분석 결과에 기반한 3대 핵심 추진 전략입니다.")
        
        cards = st.columns(3)
        with cards[0]:
            st.success("##### 전략A: 콘텐츠 성지순례")
            st.write("**근거:** SNS 언급량 1위 (15.5만 건)")
            st.write("**대상:** 부산, 수원, 전주")
            st.write("**방안:** K-pop 공연 및 드라마 촬영지 연계")
        with cards[1]:
            st.success("##### 전략B: 뷰티·웰니스 패키지")
            st.write("**근거:** 웰니스 업종 16.2% 점유")
            st.write("**대상:** 제주, 경주, 강릉")
            st.write("**방안:** 한방 스파 및 럭셔리 스테이 결합")
        with cards[2]:
            st.success("##### 전략C: 지방 유도 프로모션")
            st.write("**근거:** 결제 격차 46%p 발생")
            st.write("**대상:** 인천, 수원, 춘천")
            st.write("**방안:** 간편결제 혜택 및 교통권 결합")

        st.markdown("---")
        
        col_r, col_t = st.columns([6, 4])
        with col_r:
            st.subheader("🎨 콘텐츠 선호도 레이더 차트")
            cats = ['드라마', '음악', '음식', '패션', '뷰티']
            fig_r = go.Figure()
            fig_r.add_trace(go.Scatterpolar(r=[80, 90, 85, 70, 75], theta=cats, fill='toself', name='여성', line_color="red"))
            fig_r.add_trace(go.Scatterpolar(r=[60, 70, 75, 50, 45], theta=cats, fill='toself', name='남성', line=dict(dash='dash', color="blue")))
            fig_r.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 100])), template="plotly_white")
            st.plotly_chart(fig_r, use_container_width=True)
            
        with col_t:
            st.subheader("📌 지역별 기회 Matrix")
            s_df = pd.DataFrame({
                "지역": ["강원", "부산", "경북", "전주", "수원"],
                "일본인 비율": ["1.27%", "9.75%", "1.86%", "측정중", "경기포함"],
                "핵심자원": ["자연·드라마", "해운대·K팝", "경주·문화", "한옥·한식", "화성·민속촌"],
                "추천전략": ["웰니스 패키지", "성지순례", "한문화 체험", "식도락 투어", "당일치기"]
            })
            st.table(s_df)
            st.download_button("📊 지역 Matrix 다운로드 (CSV)", s_df.to_csv().encode('utf-8'), "regional_matrix.csv", "text/csv")
