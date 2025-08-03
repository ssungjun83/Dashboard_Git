import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os
import re
from datetime import date, timedelta

# --- 페이지 기본 설정 ---
st.set_page_config(layout="wide", page_title="지능형 생산 대시보드 V105 (차트 축 자동 범위 최적화 V4)", page_icon="👑")

# --- 데이터 로딩 및 캐싱 ---
@st.cache_data
def load_all_data():
    """
    [V90 수정] 파일 로딩의 안정성을 극대화하기 위해 '정규화' 로직을 도입했습니다.
    파일 이름에서 괄호 '()'와 공백을 모두 제거한 후 키워드와 비교하여, 눈에 보이지 않는 문자나 특수문자로 인해 파일 검색이 실패하는 문제를 원천적으로 방지합니다.
    이 로직은 모든 파일(.xlsx, .xls) 검색에 적용됩니다.
    """
    data_frames = {}
    keywords = {
        'target': '목표달성율', 
        'yield': '수율', 
        'utilization': '가동률', 
        'low_util': '저가동설비',
        'defect': ('불량실적현황', '최적화')
    }
    
    current_directory = '.'
    all_files_in_dir = os.listdir(current_directory)

    for key, keyword_info in keywords.items():
        try:
            relevant_files = []
            for f in all_files_in_dir:
                filename_without_ext, ext = os.path.splitext(f)
                
                if ext.lower() not in ['.xlsx', '.xls']:
                    continue
                
                normalized_name = filename_without_ext.replace("(", "").replace(")", "").replace(" ", "")

                if key == 'defect':
                    kw_base, kw_opt = keyword_info
                    if kw_base in normalized_name and kw_opt in normalized_name:
                        relevant_files.append(f)
                else:
                    if keyword_info in normalized_name:
                        relevant_files.append(f)

            if relevant_files:
                latest_file = max(relevant_files, key=lambda f: os.path.getmtime(os.path.join(current_directory, f)))
                file_path = os.path.join(current_directory, latest_file)
                df = pd.read_excel(file_path, engine=None)
                
                for col in df.columns:
                    if df[col].dtype == 'object' and ('%' in str(df[col].iloc[0]) if not df[col].empty and df[col].iloc[0] is not None else False):
                        df[col] = df[col].astype(str).str.replace('%', '', regex=False).str.strip()
                        df[col] = pd.to_numeric(df[col], errors='coerce')
                
                if key == 'defect':
                    cols = pd.Series(df.columns)
                    for dup in cols[cols.duplicated()].unique():
                        cols[cols[cols == dup].index.values.tolist()] = [f"{dup}_{i}" if i != 0 else dup for i in range(sum(cols == dup))]
                    df.columns = cols
                    
                    rename_dict = {}
                    if '불량수량(유형별)' in df.columns: rename_dict['불량수량(유형별)'] = '유형별_불량수량'
                    if '불량수량(전체)' in df.columns: rename_dict['불량수량(전체)'] = '총_불량수량'
                    elif '불량수량' in df.columns and '불량수량_1' in df.columns:
                        rename_dict['불량수량'] = '총_불량수량'
                        rename_dict['불량수량_1'] = '유형별_불량수량'
                    df = df.rename(columns=rename_dict)

                data_frames[key] = (df, latest_file)
            else:
                 data_frames[key] = (pd.DataFrame(), None)
        except Exception:
            data_frames[key] = (pd.DataFrame(), None)
    return data_frames

# --- AI 분석 엔진 ---
def analyze_target_data(df): return "#### AI Analyst 브리핑\n'양품 기반 달성률'을 기준으로 공장/공정별 성과를 비교하고, 목표 대비 **양품 수량**의 차이가 큰 항목을 확인하여 품질 및 생산성 개선 포인트를 동시에 도출해야 합니다."
def analyze_yield_data(df): return "#### AI Analyst 브리핑\n'수율'은 품질 경쟁력의 핵심 지표입니다. 수율이 낮은 공정/품명을 식별하고, 생산량 대비 양품 수량의 차이를 분석하여 원인을 개선해야 합니다."
def analyze_utilization_data(df): return "#### AI Analyst 브리핑\n'가동률'은 생산 효율성을 나타냅니다. 이론적인 생산 능력(CAPA)과 실제 생산량의 차이를 분석하여, 유휴 시간 및 비가동 손실을 최소화해야 합니다."
def analyze_low_utilization_data(df):
    if df is None or df.empty: return "#### AI Analyst 브리핑\n\n기준 미달인 저가동 설비가 없어, 모든 설비가 효율적으로 운영되고 있습니다."
    return "#### AI Analyst 브리핑\n'저가동 설비'는 고정비 부담 요인입니다. 가동률이 기준에 미달하는 설비의 현황을 파악하고, 유휴 자산의 효율적인 활용 방안(재배치/매각 등)을 검토해야 합니다."
def analyze_defect_data(df): return "#### AI Analyst 브리핑\n'파레토 분석'은 '80/20 법칙'에 기반하여, 소수의 핵심 불량 원인이 전체 문제의 대부분을 차지한다고 봅니다. 차트의 왼쪽에서부터 가장 큰 비중을 차지하는 불량 유형에 집중하여 개선 활동을 펼치면, 최소의 노력으로 최대의 품질 개선 효과를 얻을 수 있습니다."

# --- Helper Functions ---
PROCESS_MASTER_ORDER = ['[10] 사출조립', '[20] 분리', '[45] 하이드레이션/전면검사', '[55] 접착/멸균', '[80] 누수/규격검사']

def normalize_process_codes(df):
    """공정 컬럼의 값을 표준화하고, 컬럼명을 '공정코드'로 통일하며, 안정성을 높입니다."""
    process_col_name = None
    if '공정코드' in df.columns: process_col_name = '공정코드'
    elif '공정' in df.columns: process_col_name = '공정'
    else: return df
    df[process_col_name] = df[process_col_name].astype(str).str.strip()
    process_map = {re.search(r'\[(\d+)\]', name).group(1): name for name in PROCESS_MASTER_ORDER}
    def map_process(process_name):
        if not isinstance(process_name, str): return process_name
        match = re.search(r'\[(\d+)\]', process_name)
        return process_map.get(match.group(1), process_name) if match else process_name
    df[process_col_name] = df[process_col_name].apply(map_process)
    if process_col_name == '공정': df = df.rename(columns={'공정': '공정코드'})
    return df

def get_process_order(df, col_name='공정코드'):
    if col_name not in df.columns: return []
    processes_in_df = df[col_name].unique()
    return [p for p in PROCESS_MASTER_ORDER if p in processes_in_df]

def add_date_column(df, date_col_name=None):
    """다양한 날짜 컬럼명을 'date'로 통일하여 새 컬럼을 추가합니다."""
    if 'date' in df.columns:
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        return df
    date_candidates = [date_col_name, '생산일자', '일자', '기간'] if date_col_name else ['생산일자', '일자', '기간']
    found_col = next((col for col in date_candidates if col in df.columns), None)
    if found_col:
        if found_col == '기간': df['date'] = pd.to_datetime(df[found_col].astype(str).str.split(' ~ ').str[0], errors='coerce')
        else: df['date'] = pd.to_datetime(df[found_col], errors='coerce')
    else: df['date'] = pd.NaT
    return df

def get_resampled_data(df, agg_level, metrics_to_sum, group_by_cols=['period', '공장', '공정코드']):
    if df.empty or 'date' not in df.columns or df['date'].isnull().all(): return pd.DataFrame()
    df_copy = df.copy().dropna(subset=['date'])
    if agg_level == '일별':
        df_copy['period'] = df_copy['date'].dt.strftime('%Y-%m-%d')
    elif agg_level == '주간별':
        start_of_week = df_copy['date'] - pd.to_timedelta(df_copy['date'].dt.dayofweek, unit='d')
        end_of_week = start_of_week + pd.to_timedelta(6, unit='d')
        df_copy['period'] = start_of_week.dt.strftime('%Y-%m-%d') + ' ~ ' + end_of_week.dt.strftime('%Y-%m-%d')
    elif agg_level == '월별':
        df_copy['period'] = df_copy['date'].dt.strftime('%Y-%m')
    elif agg_level == '분기별':
        df_copy['period'] = df_copy['date'].dt.year.astype(str) + '년 ' + df_copy['date'].dt.quarter.astype(str) + '분기'
    elif agg_level == '반기별':
        df_copy['period'] = df_copy['date'].dt.year.astype(str) + '년 ' + df_copy['date'].dt.month.apply(lambda m: '상반기' if m <= 6 else '하반기')
    elif agg_level == '년도별':
        df_copy['period'] = df_copy['date'].dt.strftime('%Y')
    else:
        df_copy['period'] = df_copy['date'].dt.strftime('%Y-%m-%d')
        
    valid_group_by_cols = [col for col in group_by_cols if col in df_copy.columns or col == 'period']
    agg_dict = {metric: 'sum' for metric in metrics_to_sum if metric in df_copy.columns}
    if not agg_dict:
        if 'period' not in df_copy.columns: return pd.DataFrame(columns=valid_group_by_cols)
        return df_copy[valid_group_by_cols].drop_duplicates()
    return df_copy.groupby(valid_group_by_cols).agg(agg_dict).reset_index()

def generate_summary_text(df, agg_level, factory_name="전체"):
    agg_map = {'일별': '일', '주간별': '주', '월별': '월', '분기별': '분기', '반기별': '반기', '년도별': '년'}
    period_text = agg_map.get(agg_level, '기간')
    title_prefix = f"{factory_name} " if factory_name != "전체" else ""
    if df.empty or len(df) < 2: return f"""<div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; margin-bottom: 20px; font-family: 'Malgun Gothic', sans-serif; background-color: #f9f9f9; line-height: 1.6;"><h4 style="margin-top:0; color: #1E88E5; font-size: 1.3em;">{title_prefix}AI Analyst 종합 분석 브리핑</h4><p style="font-size: 1.1em;">분석할 데이터가 부족하여 추이 분석을 제공할 수 없습니다. 최소 2개 이상의 {period_text}치 데이터를 선택해주세요.</p></div>"""
    df = df.copy(); start_period = df['period'].iloc[0]; end_period = df['period'].iloc[-1]; total_prod = df['총_생산수량'].sum(); avg_prod = df['총_생산수량'].mean(); max_prod_row = df.loc[df['총_생산수량'].idxmax()]; min_prod_row = df.loc[df['총_생산수량'].idxmin()]
    first_prod = df['총_생산수량'].iloc[0]; last_prod = df['총_생산수량'].iloc[-1]; prod_change = last_prod - first_prod; prod_change_pct = (prod_change / first_prod * 100) if first_prod != 0 else 0; prod_trend_text = "증가" if prod_change > 0 else "감소" if prod_change < 0 else "유지"
    avg_yield = df['종합수율(%)'].mean(); max_yield_row = df.loc[df['종합수율(%)'].idxmax()]; min_yield_row = df.loc[df['종합수율(%)'].idxmin()]; first_yield = df['종합수율(%)'].iloc[0]; last_yield = df['종합수율(%)'].iloc[-1]; yield_change = last_yield - first_yield; yield_trend_text = "개선" if yield_change > 0 else "하락" if yield_change < 0 else "유지"
    insight_text = ""
    if len(df) >= 3:
        correlation = df['총_생산수량'].corr(df['종합수율(%)']); max_yield_row_insight = df.loc[df['종합수율(%)'].idxmax()]; max_prod_row_insight = df.loc[df['총_생산수량'].idxmax()]
        if correlation > 0.5: insight_text = (f"<strong>긍정적 신호:</strong> 생산량과 수율 간에 강한 양의 상관관계(상관계수: {correlation:.2f})가 나타났습니다. 특히 생산량과 수율이 모두 정점에 달했던 <strong>{max_prod_row_insight['period']}</strong> 또는 <strong>{max_yield_row_insight['period']}</strong>의 성공 요인을 분석하여, 이를 전체 공정에 확산시킬 필요가 있습니다.")
        elif correlation < -0.5: insight_text = (f"<strong>주의 필요:</strong> 생산량과 수율 간에 강한 음의 상관관계(상관계수: {correlation:.2f})가 발견되었습니다. 이는 생산량을 늘릴수록 수율이 떨어지는 경향을 의미합니다. 생산량이 가장 많았던 <strong>{max_prod_row_insight['period']}</strong>의 수율(<strong>{max_prod_row_insight['종합수율(%)']:.2f}%</strong>)이 평균 이하인 점을 주목하고, 해당 기간의 불량 원인을 집중 분석해야 합니다.")
        else: insight_text = (f"<strong>독립적 관계:</strong> 생산량과 수율 간의 뚜렷한 상관관계(상관계수: {correlation:.2f})는 보이지 않습니다. 수율이 가장 높았던 <strong>{max_yield_row_insight['period']}</strong>(<strong>{max_yield_row_insight['종합수율(%)']:.2f}%</strong>)의 사례를 분석하여, 수율을 높일 수 있는 독립적인 개선 방안을 도출해야 합니다.")
    else: insight_text = (f"<strong>{df.loc[df['총_생산수량'].idxmax()]['period']}</strong>에 생산량이 정점을 찍었을 때, 수율은 <strong>{df.loc[df['총_생산수량'].idxmax()]['종합수율(%)']:.2f}%</strong>를 기록했습니다. {agg_level} 생산량과 수율의 관계를 지속적으로 모니터링하여 최적의 생산 조건을 찾아야 합니다.")
    summary = f"""
<div style="border: 1px solid #e0e0e0; border-radius: 8px; padding: 20px; margin-bottom: 20px; font-family: 'Malgun Gothic', sans-serif; background-color: #f9f9f9; line-height: 1.6;">
    <h4 style="margin-top:0; color: #1E88E5; font-size: 1.3em;">{title_prefix}AI Analyst 종합 분석 브리핑 ({agg_level})</h4>
    <p style="font-size: 1.0em;"><strong>분석 기간:</strong> {start_period} ~ {end_period}</p>
    <ul style="list-style-type: none; padding-left: 0; font-size: 1.1em;">
        <li style="margin-bottom: 10px;">
            <span style="font-size: 1.2em; vertical-align: middle;">📈</span> <strong>생산 실적:</strong>
            분석 기간 동안 총 <strong style="color: #004D40;">{total_prod:,.0f}개</strong>를 생산했으며, {period_text} 평균 <strong style="color: #004D40;">{avg_prod:,.0f}개</strong>를 생산했습니다.
            생산량은 <strong style="color: #1565C0;">{max_prod_row['period']}</strong>에 <strong style="color: #1565C0;">{max_prod_row['총_생산수량']:,.0f}개</strong>로 최고치를,
            <strong style="color: #C62828;">{min_prod_row['period']}</strong>에 <strong style="color: #C62828;">{min_prod_row['총_생산수량']:,.0f}개</strong>로 최저치를 기록했습니다.
            기간 전체적으로 생산량은 <strong style="color: {'#1E88E5' if prod_change > 0 else '#E53935'};">{abs(prod_change_pct):.2f}% {prod_trend_text}</strong>하는 추세를 보였습니다.
        </li>
        <li>
            <span style="font-size: 1.2em; vertical-align: middle;">⚙️</span> <strong>종합 수율:</strong>
            기간 내 {period_text} 평균 종합 수율은 <strong style="color: #004D40;">{avg_yield:.2f}%</strong> 입니다.
            수율은 <strong style="color: #1565C0;">{max_yield_row['period']}</strong>에 <strong style="color: #1565C0;">{max_yield_row['종합수율(%)']:.2f}%</strong>로 가장 높았고,
            <strong style="color: #C62828;">{min_yield_row['period']}</strong>에 <strong style="color: #C62828;">{min_yield_row['종합수율(%)']:.2f}%</strong>로 가장 낮았습니다.
            전반적으로 수율은 <strong style="color: {'#1E88E5' if yield_change > 0 else '#E53935'};">{yield_trend_text}</strong>되었습니다.
        </li>
    </ul>
    <div style="margin-top: 15px; padding-top: 15px; border-top: 1px solid #ddd;">
        <p style="font-size: 1.1em;"><strong><span style="font-size: 1.2em; vertical-align: middle;">💡</span> 핵심 인사이트:</strong> {insight_text}</p>
    </div>
</div>
"""
    return summary

def plot_pareto_chart(df, title, defect_qty_col='유형별_불량수량'):
    if df.empty or defect_qty_col not in df.columns: 
        st.info("차트를 그릴 데이터가 없습니다.")
        return
    df_agg = df.groupby('불량명')[defect_qty_col].sum().reset_index()
    df_agg = df_agg.sort_values(by=defect_qty_col, ascending=False)
    df_agg = df_agg[df_agg[defect_qty_col] > 0] 
    if df_agg.empty: 
        st.info("선택된 항목에 보고된 불량이 없습니다.")
        return
    df_agg['누적합계'] = df_agg[defect_qty_col].cumsum()
    df_agg['누적비율'] = (df_agg['누적합계'] / df_agg[defect_qty_col].sum()) * 100
    
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    
    fig.add_trace(go.Bar(
        x=df_agg['불량명'], 
        y=df_agg[defect_qty_col], 
        name='불량 수량', 
        text=df_agg[defect_qty_col], 
        texttemplate='%{text:,.0f}', 
        textposition='outside',
        textfont=dict(size=18, family="Arial, sans-serif", color="black")
    ), secondary_y=False)
    
    fig.add_trace(go.Scatter(
        x=df_agg['불량명'], 
        y=df_agg['누적비율'], 
        name='누적 비율', 
        mode='lines+markers+text',
        text=df_agg['누적비율'], 
        texttemplate='%{text:.1f}%', 
        textposition='top center',
        textfont=dict(size=16, color='black') 
    ), secondary_y=True)
    
    fig.update_layout(height=600, title_text=f'<b>{title}</b>', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
    fig.update_yaxes(title_text="<b>불량 수량 (개)</b>", secondary_y=False)
    fig.update_yaxes(title_text="<b>누적 비율 (%)</b>", secondary_y=True, range=[0, 105])
    fig.update_xaxes(title_text="<b>불량 유형</b>")
    st.plotly_chart(fig, use_container_width=True)

def reset_filters(min_data_date, max_data_date):
    """Callback function to reset date range to the full data range and agg_level to '월별'."""
    st.session_state.date_range = (min_data_date, max_data_date)
    st.session_state.agg_level = '월별'

# --- 대시보드 UI 시작 ---
st.title("지능형 생산 대시보드 V105 (차트 축 자동 범위 최적화 V4)")

all_data = load_all_data()
df_target_orig, target_filename = all_data.get('target', (pd.DataFrame(), None)); df_yield_orig, yield_filename = all_data.get('yield', (pd.DataFrame(), None)); df_utilization_orig, util_filename = all_data.get('utilization', (pd.DataFrame(), None)); df_low_util_orig, low_util_filename = all_data.get('low_util', (pd.DataFrame(), None)); df_defect_orig, defect_filename = all_data.get('defect', (pd.DataFrame(), None))

if not df_target_orig.empty: df_target_orig = normalize_process_codes(add_date_column(df_target_orig))
if not df_yield_orig.empty: df_yield_orig = normalize_process_codes(add_date_column(df_yield_orig))
if not df_utilization_orig.empty: df_utilization_orig = normalize_process_codes(add_date_column(df_utilization_orig))
if not df_defect_orig.empty: df_defect_orig = normalize_process_codes(add_date_column(df_defect_orig))

if 'date_range' not in st.session_state or 'agg_level' not in st.session_state:
    all_dfs = [df_target_orig, df_yield_orig, df_utilization_orig, df_defect_orig]
    all_dates = pd.concat([d['date'] for d in all_dfs if d is not None and not d.empty and 'date' in d.columns]).dropna()
    min_date_global, max_date_global = (all_dates.min().date(), all_dates.max().date()) if not all_dates.empty else (date.today(), date.today())
    if 'date_range' not in st.session_state: st.session_state.date_range = (min_date_global, max_date_global)
    if 'agg_level' not in st.session_state: st.session_state.agg_level = '월별'

st.sidebar.header("로딩된 파일 정보"); st.sidebar.info(f"목표: {target_filename}" if target_filename else "파일 없음"); st.sidebar.info(f"수율: {yield_filename}" if yield_filename else "파일 없음"); st.sidebar.info(f"가동률: {util_filename}" if util_filename else "파일 없음"); st.sidebar.info(f"저가동: {low_util_filename}" if low_util_filename else "파일 없음"); st.sidebar.info(f"불량: {defect_filename}" if defect_filename else "파일 없음")

tab_list = ["종합 분석", "목표 달성률", "수율 분석", "불량유형별 분석", "가동률 분석", "저가동 설비"]
selected_tab = st.radio("메인 네비게이션", tab_list, key='main_tab_selector', horizontal=True, label_visibility='collapsed')

def create_shared_filter_controls(df_for_current_tab):
    """
    모든 탭에서 공유되는 필터 컨트롤을 생성하고 필터링된 데이터프레임을 반환합니다.
    """
    all_dfs = [df_target_orig, df_yield_orig, df_utilization_orig, df_defect_orig]
    all_dates = pd.concat([d['date'] for d in all_dfs if d is not None and not d.empty and 'date' in d.columns]).dropna()
    min_date_global, max_date_global = (all_dates.min().date(), all_dates.max().date()) if not all_dates.empty else (date(2000, 1, 1), date.today())

    header_cols = st.columns([1, 1])
    with header_cols[0]:
        header_title = selected_tab
        if "분석" not in selected_tab: header_title = f"{selected_tab} 분석"
        st.header(header_title, anchor=False)

    filter_cols = st.columns([6, 1, 3.5])
    with filter_cols[0]:
        st.date_input("조회할 기간을 선택하세요", min_value=min_date_global, max_value=max_date_global, key='date_range')
    with filter_cols[1]:
        st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
        st.button("기간 초기화", on_click=reset_filters, args=(min_date_global, max_date_global), help="조회 기간을 데이터의 전체 기간으로, 집계 기준을 '월별'로 초기화합니다.")
    with filter_cols[2]:
        st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
        st.radio("집계 기준", options=['일별', '주간별', '월별', '분기별', '반기별', '년도별'], key='agg_level', horizontal=True)

    date_range_value = st.session_state.get('date_range')
    agg_level = st.session_state.get('agg_level', '월별')

    if isinstance(date_range_value, (list, tuple)) and len(date_range_value) == 2:
        start_date, end_date = date_range_value
    else:
        start_date, end_date = min_date_global, max_date_global
    
    final_start_date = max(start_date, min_date_global)
    final_end_date = min(end_date, max_date_global)
    
    with header_cols[1]:
        st.markdown(f"<p style='text-align: right; margin-top: 1.2rem; font-size: 1.1rem; color: grey;'>({final_start_date.strftime('%Y-%m-%d')} ~ {final_end_date.strftime('%Y-%m-%d')})</p>", unsafe_allow_html=True)
    
    if df_for_current_tab.empty or 'date' not in df_for_current_tab.columns or df_for_current_tab['date'].isnull().all():
        return pd.DataFrame(), final_start_date, final_end_date, agg_level
        
    mask = (df_for_current_tab['date'].dt.date >= final_start_date) & (df_for_current_tab['date'].dt.date <= final_end_date)
    return df_for_current_tab[mask].copy(), final_start_date, final_end_date, agg_level

def aggregate_overall_data(df, analysis_type):
    if df.empty: return pd.DataFrame()
    group_cols = ['공장', '공정코드']
    metrics_map = {'target': {'sums': ['목표_총_생산량', '총_생산수량'], 'rate': '달성률(%)'}, 'yield': {'sums': ['총_생산수량', '총_양품수량'], 'rate': '평균_수율'}, 'utilization': {'sums': ['총_생산수량', '이론상_총_생산량'], 'rate': '평균_가동률'}}
    metrics = metrics_map.get(analysis_type);
    if not metrics: return pd.DataFrame()
    agg_dict = {col: 'sum' for col in metrics['sums'] if col in df.columns};
    if not agg_dict: return pd.DataFrame()
    agg_df = df.groupby(group_cols).agg(agg_dict).reset_index()
    rate_name, sums = metrics['rate'], metrics['sums']
    c1, c2 = sums if analysis_type != 'utilization' else (sums[1], sums[0])
    with pd.option_context('mode.use_inf_as_na', True): agg_df[rate_name] = (100 * agg_df[c2] / agg_df[c1]).fillna(0)
    return agg_df

def plot_horizontal_bar_chart_all_processes(df, analysis_info, all_factories, all_processes):
    rate_col, y_axis_title, chart_title = analysis_info['rate_col'], analysis_info['y_axis_title'], analysis_info['chart_title']
    all_combinations = pd.DataFrame([(f, p) for f in all_factories for p in all_processes], columns=['공장', '공정코드'])
    df_complete = pd.merge(all_combinations, df, on=['공장', '공정코드'], how='left')
    df_complete[rate_col] = df_complete[rate_col].fillna(0)
    st.divider(); st.subheader("공장/공정별 현황 (전체 기간 집계)")
    df_complete['공정코드'] = pd.Categorical(df_complete['공정코드'], categories=all_processes, ordered=True)
    df_complete = df_complete.sort_values(by=['공장', '공정코드']); category_orders = {'공정코드': all_processes}
    fig = px.bar(df_complete, x=rate_col, y='공정코드', color='공장', text=rate_col, title=f'<b>{chart_title}</b>', orientation='h', facet_row="공장", height=600, facet_row_spacing=0.05, category_orders=category_orders)
    fig.update_traces(texttemplate='%{text:.2f}%', textposition='auto'); fig.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1])); fig.update_yaxes(title=y_axis_title)
    st.plotly_chart(fig, use_container_width=True)

# --- 탭별 UI 구현 ---
if selected_tab == "목표 달성률":
    if df_target_orig.empty or df_yield_orig.empty: st.info("해당 분석을 위해서는 '목표달성율'과 '수율' 데이터가 모두 필요합니다.")
    else:
        df_target_filtered, start_date, end_date, agg_level = create_shared_filter_controls(df_target_orig)
        if df_target_filtered.empty: st.info("선택된 기간에 목표 데이터가 없습니다.")
        else:
            mask_yield = (df_yield_orig['date'].dt.date >= start_date) & (df_yield_orig['date'].dt.date <= end_date); df_yield_filtered = df_yield_orig.loc[mask_yield].copy()
            if df_yield_filtered.empty: st.info("선택된 기간에 수율 데이터가 없어, 양품 기반 달성률을 계산할 수 없습니다.")
            else:
                key_cols = ['date', '공장', '공정코드']; target_agg_day = df_target_filtered.groupby(key_cols).agg(목표_총_생산량=('목표_총_생산량', 'sum')).reset_index(); yield_agg_day = df_yield_filtered.groupby(key_cols).agg(총_생산수량=('총_생산수량', 'sum'), 총_양품수량=('총_양품수량', 'sum')).reset_index()
                df_merged = pd.merge(target_agg_day, yield_agg_day, on=key_cols, how='outer'); df_merged.fillna({'총_양품수량': 0, '총_생산수량': 0, '목표_총_생산량': 0}, inplace=True); main_col, side_col = st.columns([2.8, 1])
                with main_col:
                    st.subheader("핵심 지표 요약 (완제품 제조 기준, 양품 기반 달성률)"); df_kpi_base = df_merged[df_merged['공정코드'] == '[80] 누수/규격검사']
                    if not df_kpi_base.empty:
                        df_kpi_agg_factory = df_kpi_base.groupby('공장').agg(목표_총_생산량=('목표_총_생산량', 'sum'), 총_양품수량=('총_양품수량', 'sum')).reset_index()
                        with pd.option_context('mode.use_inf_as_na', True): df_kpi_agg_factory['달성률(%)'] = (100 * df_kpi_agg_factory['총_양품수량'] / df_kpi_agg_factory['목표_총_생산량']).fillna(0)
                        target_kpi, good_kpi = df_kpi_agg_factory['목표_총_생산량'].sum(), df_kpi_agg_factory['총_양품수량'].sum(); rate_kpi = (good_kpi / target_kpi * 100) if target_kpi > 0 else 0
                        kpi1, kpi2, kpi3 = st.columns(3); kpi1.metric("완제품 목표", f"{target_kpi:,.0f} 개"); kpi2.metric("완제품 양품 실적", f"{good_kpi:,.0f} 개"); kpi3.metric("완제품 달성률", f"{rate_kpi:.2f} %")
                        st.divider(); st.markdown("##### 공장별 최종 완제품 달성률 (양품 기준)"); factory_kpi_cols = st.columns(len(df_kpi_agg_factory) or [1])
                        for i, row in df_kpi_agg_factory.iterrows():
                            with factory_kpi_cols[i]: st.metric(label=row['공장'], value=f"{row['달성률(%)']:.2f}%"); st.markdown(f"<p style='font-size:0.8rem;color:grey;margin-top:-8px;'>목표:{row['목표_총_생산량']:,.0f}<br>양품실적:{row['총_양품수량']:,.0f}</p>", unsafe_allow_html=True)
                    st.divider(); st.subheader(f"{agg_level} 완제품 달성률 추이 (양품 기준)"); df_resampled = get_resampled_data(df_merged, agg_level, ['목표_총_생산량', '총_양품수량']); df_trend = df_resampled[df_resampled['공정코드'] == '[80] 누수/규격검사'].copy()
                    if not df_trend.empty:
                        with pd.option_context('mode.use_inf_as_na', True): df_trend['달성률(%)'] = (100 * df_trend['총_양품수량'] / df_trend['목표_총_생산량']).fillna(0)
                        fig_trend = px.line(df_trend.sort_values('period'), x='period', y='달성률(%)', color='공장', title=f'<b>{agg_level} 완제품 제조 달성률 추이 (양품 기준)</b>', markers=True, text='달성률(%)'); fig_trend.update_traces(texttemplate='%{text:.2f}%', textposition='top center', textfont=dict(size=16, color='black')); fig_trend.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(df_trend['period'].unique())); st.plotly_chart(fig_trend, use_container_width=True)
                    df_total_agg = df_merged.groupby(['공장', '공정코드']).agg(목표_총_생산량=('목표_총_생산량', 'sum'), 총_양품수량=('총_양품수량', 'sum')).reset_index()
                    with pd.option_context('mode.use_inf_as_na', True): df_total_agg['달성률(%)'] = (100 * df_total_agg['총_양품수량'] / df_total_agg['목표_총_생산량']).fillna(0)
                    df_total_agg = df_total_agg[df_total_agg['목표_총_생산량'] > 0]; st.divider(); st.subheader("공장/공정별 현황 (전체 기간 집계)")
                    chart_process_order = get_process_order(df_total_agg)
                    df_total_agg['공정코드'] = pd.Categorical(df_total_agg['공정코드'], categories=chart_process_order, ordered=True); df_total_agg = df_total_agg.sort_values(by=['공장', '공정코드']); category_orders = {'공정코드': chart_process_order}
                    fig_bar = px.bar(df_total_agg, x='달성률(%)', y='공정코드', color='공장', text='달성률(%)', title='<b>공장/공정별 달성률 현황 (양품 기준)</b>', orientation='h', facet_row="공장", height=600, facet_row_spacing=0.05, category_orders=category_orders)
                    fig_bar.update_traces(texttemplate='%{text:.2f}%', textposition='auto'); fig_bar.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1])); fig_bar.update_yaxes(title="공정"); st.plotly_chart(fig_bar, use_container_width=True)
                with side_col:
                    st.markdown(analyze_target_data(df_merged)); st.divider(); st.subheader("데이터 원본 (일별 집계)"); df_display = df_merged.copy();
                    with pd.option_context('mode.use_inf_as_na', True): df_display['달성률(%)'] = (100 * df_display['총_양품수량'] / df_display['목표_총_생산량']).fillna(0)
                    df_display = df_display.rename(columns={'date': '일자', '목표_총_생산량': '목표 생산량', '총_생산수량': '총 생산량', '총_양품수량': '총 양품수량'}); st.dataframe(df_display[['일자', '공장', '공정코드', '목표 생산량', '총 생산량', '총 양품수량', '달성률(%)']].sort_values(by=['일자', '공장', '공정코드']), use_container_width=True, height=500)

elif selected_tab == "수율 분석":
    df_filtered, _, _, agg_level = create_shared_filter_controls(df_yield_orig)
    if not df_filtered.empty:
        main_col, side_col = st.columns([2.8, 1])
        with main_col:
            # --- 공장별 종합 수율 추이 ---
            df_resampled_factory = get_resampled_data(df_filtered, agg_level, ['총_생산수량', '총_양품수량'], group_by_cols=['period', '공장', '공정코드'])
            if not df_resampled_factory.empty:
                st.subheader(f"{agg_level} 공장별 종합 수율 추이")
                with pd.option_context('mode.use_inf_as_na', True): df_resampled_factory['개별수율'] = (df_resampled_factory['총_양품수량'] / df_resampled_factory['총_생산수량']).fillna(1.0)
                factory_yield_trend = df_resampled_factory.groupby(['period', '공장'])['개별수율'].prod().reset_index()
                factory_yield_trend['종합수율(%)'] = factory_yield_trend.pop('개별수율') * 100
                fig_factory_trend = px.line(factory_yield_trend.sort_values('period'), x='period', y='종합수율(%)', color='공장', title=f'<b>{agg_level} 공장별 종합 수율 추이</b>', markers=True, text='종합수율(%)')
                fig_factory_trend.update_traces(texttemplate='%{text:.2f}%', textposition='top center', textfont=dict(size=16, color='black'))
                fig_factory_trend.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(factory_yield_trend['period'].unique()))
                st.plotly_chart(fig_factory_trend, use_container_width=True)

            st.divider()
            
            # --- 제품군별 종합 수율 추이 ---
            st.subheader(f"{agg_level} 제품군별 종합 수율 추이")
            
            # 공장 선택 필터
            all_factories = ['전체'] + sorted(df_filtered['공장'].unique())
            selected_factory = st.selectbox(
                "공장 선택", 
                options=all_factories, 
                key="yield_factory_select",
                help="분석할 공장을 선택합니다. '전체' 선택 시 모든 공장의 데이터를 종합하여 분석합니다."
            )

            # 선택된 공장에 따라 데이터 필터링
            if selected_factory == '전체':
                df_yield_factory_filtered = df_filtered.copy()
            else:
                df_yield_factory_filtered = df_filtered[df_filtered['공장'] == selected_factory].copy()
            
            df_resampled_product = get_resampled_data(df_yield_factory_filtered, agg_level, ['총_생산수량', '총_양품수량'], group_by_cols=['period', '신규분류요약', '공정코드'])

            if not df_resampled_product.empty and '신규분류요약' in df_resampled_product.columns:
                with pd.option_context('mode.use_inf_as_na', True): 
                    df_resampled_product['개별수율'] = (df_resampled_product['총_양품수량'] / df_resampled_product['총_생산수량']).fillna(1.0)
                
                product_yield_trend = df_resampled_product.groupby(['period', '신규분류요약'])['개별수율'].prod().reset_index()
                product_yield_trend = product_yield_trend.rename(columns={'개별수율': '종합수율(%)'})
                product_yield_trend['종합수율(%)'] *= 100
                
                all_product_groups = sorted(df_resampled_product['신규분류요약'].dropna().unique())

                if not all_product_groups:
                    st.info("선택된 공장에 제품군 데이터가 없습니다.")
                else:
                    for group in all_product_groups:
                        if f"product_group_{group}" not in st.session_state: 
                            st.session_state[f"product_group_{group}"] = True
                    
                    st.markdown("##### 표시할 제품군 선택")
                    btn_cols = st.columns(8)
                    with btn_cols[0]:
                        if st.button("제품군 전체 선택", key="select_all_products_yield", use_container_width=True):
                            for group in all_product_groups: st.session_state[f"product_group_{group}"] = True
                            st.rerun()
                    with btn_cols[1]:
                        if st.button("제품군 전체 해제", key="deselect_all_products_yield", use_container_width=True):
                            for group in all_product_groups: st.session_state[f"product_group_{group}"] = False
                            st.rerun()
                    
                    st.write("")
                    num_cols = 5
                    cols = st.columns(num_cols)
                    selected_product_groups = []
                    for i, group in enumerate(all_product_groups):
                        with cols[i % num_cols]:
                            if st.checkbox(group, key=f"product_group_{group}"):
                                selected_product_groups.append(group)
                    
                    combine_yield = st.checkbox("선택항목 합쳐서 보기", key="combine_product_yield", help="선택한 제품군들의 실적을 합산하여 단일 종합 수율 추이를 분석합니다.")

                    if selected_product_groups:
                        if combine_yield:
                            df_filtered_for_combine = df_resampled_product[df_resampled_product['신규분류요약'].isin(selected_product_groups)]
                            df_combined = df_filtered_for_combine.groupby(['period', '공정코드']).agg(총_생산수량=('총_생산수량', 'sum'), 총_양품수량=('총_양품수량', 'sum')).reset_index()
                            with pd.option_context('mode.use_inf_as_na', True): 
                                df_combined['개별수율'] = (df_combined['총_양품수량'] / df_combined['총_생산수량']).fillna(1.0)
                            
                            df_to_plot = df_combined.groupby('period')['개별수율'].prod().reset_index()
                            df_to_plot = df_to_plot.rename(columns={'개별수율': '종합수율(%)'})
                            df_to_plot['종합수율(%)'] *= 100
                            
                            if not df_to_plot.empty:
                                fig_product_trend = px.line(df_to_plot.sort_values('period'), x='period', y='종합수율(%)', title=f'<b>{agg_level} 선택 제품군 통합 수율 추이 ({selected_factory})</b>', markers=True, text='종합수율(%)')
                                fig_product_trend.update_traces(texttemplate='%{text:.2f}%', textposition='top center', textfont=dict(size=16, color='black'))
                                fig_product_trend.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(df_to_plot['period'].unique()))
                                st.plotly_chart(fig_product_trend, use_container_width=True)
                        else:
                            df_to_plot = product_yield_trend[product_yield_trend['신규분류요약'].isin(selected_product_groups)]
                            if not df_to_plot.empty:
                                fig_product_trend = px.line(df_to_plot.sort_values('period'), x='period', y='종합수율(%)', color='신규분류요약', title=f'<b>{agg_level} 제품군별 종합 수율 추이 ({selected_factory})</b>', markers=True, text='종합수율(%)')
                                fig_product_trend.update_traces(texttemplate='%{text:.2f}%', textposition='top center', textfont=dict(size=16, color='black'))
                                fig_product_trend.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(df_to_plot['period'].unique()))
                                st.plotly_chart(fig_product_trend, use_container_width=True)
                    else:
                        st.info("차트를 표시할 제품군을 선택해주세요.")

            # --- 공장/공정별 평균 수율 ---
            df_total_agg = aggregate_overall_data(df_filtered, 'yield')
            all_factories_in_period = sorted(df_filtered['공장'].unique())
            plot_horizontal_bar_chart_all_processes(df_total_agg, {'rate_col': '평균_수율', 'y_axis_title': '평균 수율', 'chart_title': '공장/공정별 평균 수율'}, all_factories_in_period, PROCESS_MASTER_ORDER)

        with side_col:
            st.markdown(analyze_yield_data(df_total_agg))
            st.divider()
            st.subheader("데이터 원본")
            st.dataframe(df_filtered, use_container_width=True, height=500)
    else:
        st.info(f"선택된 기간에 해당하는 수율 데이터가 없습니다.")

elif selected_tab == "불량유형별 분석":
    if df_defect_orig.empty:
        st.info("해당 분석을 위해서는 '불량실적현황(최적화)' 데이터가 필요합니다.")
    else:
        df_defect_filtered, _, _, agg_level = create_shared_filter_controls(df_defect_orig)

        if df_defect_filtered.empty:
            st.info("선택된 기간에 분석에 필요한 불량 데이터가 없습니다.")
        elif '생산수량' not in df_defect_filtered.columns:
            st.error("불량 데이터 파일에 '생산수량' 컬럼이 없어 불량률을 계산할 수 없습니다.")
        else:
            if '유형별_불량수량' in df_defect_filtered.columns:
                df_defect_filtered['유형별_불량수량'] = pd.to_numeric(df_defect_filtered['유형별_불량수량'], errors='coerce').fillna(0)
            
            main_col, side_col = st.columns([2.8, 1])

            with main_col:
                with st.expander("세부 필터 및 옵션", expanded=True):
                    filter_data_source = df_defect_filtered.copy()
                    filter_options_map = {
                        "공장": "공장",
                        "신규분류요약": "제품군",
                        "사출기계코드": "사출 기계",
                        "공정기계코드": "공정 기계"
                    }
                    available_filters = [k for k in filter_options_map if k in filter_data_source.columns]

                    # 최초 실행 시 모든 필터 전체 선택
                    for key in available_filters:
                        options = sorted(filter_data_source[key].dropna().unique())
                        session_key = f"ms_{key}"
                        if session_key not in st.session_state:
                            st.session_state[session_key] = options

                    # 전체 선택/해제 버튼
                    btn_cols = st.columns(2)
                    with btn_cols[0]:
                        if st.button("세부필터 전체 선택"):
                            for key in available_filters:
                                options = sorted(filter_data_source[key].dropna().unique())
                                st.session_state[f"ms_{key}"] = options
                            st.rerun()
                    with btn_cols[1]:
                        if st.button("세부필터 전체 해제"):
                            for key in available_filters:
                                st.session_state[f"ms_{key}"] = []
                            st.rerun()

                    # 동적 필터링
                    selections = {}
                    filtered_df = filter_data_source.copy()
                    for i, key in enumerate(available_filters):
                        # 앞쪽 필터 선택값에 따라 옵션 제한
                        if i > 0:
                            prev_keys = available_filters[:i]
                            for pk in prev_keys:
                                selected = st.session_state.get(f"ms_{pk}", [])
                                if selected:
                                    filtered_df = filtered_df[filtered_df[pk].isin(selected)]
                        options = sorted(filtered_df[key].dropna().unique())
                        selections[key] = st.multiselect(
                            filter_options_map[key], options, default=st.session_state.get(f"ms_{key}", options),
                            key=f"ms_{key}", label_visibility="collapsed", placeholder=filter_options_map[key]
                        )

                df_display = filtered_df.copy()
                for key, selected_values in selections.items():
                    if selected_values:
                        df_display = df_display[df_display[key].isin(selected_values)]
                
                st.markdown("---")
                st.markdown("<h6>불량 유형 필터</h6>", unsafe_allow_html=True)
                defect_options = sorted(df_display['불량명'].dropna().unique())
                if 'selected_defects' not in st.session_state: st.session_state.selected_defects = defect_options
                
                defect_btn_cols = st.columns(4)
                with defect_btn_cols[0]:
                    if st.button("불량 유형 전체 선택", use_container_width=True): st.session_state.selected_defects = defect_options
                with defect_btn_cols[1]:
                    if st.button("불량 유형 전체 해제", use_container_width=True): st.session_state.selected_defects = []
                
                st.multiselect("표시할 불량 유형 선택", options=defect_options, key='selected_defects', label_visibility="collapsed")
            
            if st.session_state.selected_defects:
                df_display = df_display[df_display['불량명'].isin(st.session_state.selected_defects)]
            else: 
                df_display = df_display[df_display['불량명'].isin([])]
            
            prod_key_cols = ['date', '공장', '신규분류요약', '사출기계코드', '공정기계코드', '생산수량']
            available_prod_key_cols = [col for col in prod_key_cols if col in df_display.columns]
            prod_data_source = df_display[available_prod_key_cols].drop_duplicates()

            st.divider()
            st.subheader("주요 불량 원인 분석 (파레토)", anchor=False)
            if df_display.empty or '유형별_불량수량' not in df_display.columns or df_display['유형별_불량수량'].sum() == 0:
                st.warning("선택된 필터 조건에 해당하는 불량 데이터가 없습니다.")
            else:
                plot_pareto_chart(df_display, title="선택된 조건의 불량유형 파레토 분석", defect_qty_col='유형별_불량수량')

            st.divider()
            st.subheader(f"{agg_level} 총 불량 수량 및 불량률 추이", anchor=False)
            total_defect_resampled = get_resampled_data(df_display, agg_level, ['유형별_불량수량'], group_by_cols=['period'])
            total_prod_resampled = get_resampled_data(prod_data_source, agg_level, ['생산수량'], group_by_cols=['period']).rename(columns={'생산수량': '총_생산수량'})
            
            if not total_defect_resampled.empty:
                combo_data = pd.merge(total_defect_resampled, total_prod_resampled, on='period', how='outer').fillna(0)
                production_for_rate = combo_data['총_생산수량'].replace(0, pd.NA)
                with pd.option_context('mode.use_inf_as_na', True):
                    combo_data['총_불량률(%)'] = (100 * combo_data['유형별_불량수량'] / production_for_rate).fillna(0)
                
                min_rate_val = combo_data['총_불량률(%)'].min()
                max_rate_val = combo_data['총_불량률(%)'].max()
                
                slider_max_bound = max(50.0, max_rate_val * 1.2)
                
                rate_range = st.slider(
                    "총 불량률(%) 축 범위 조절",
                    min_value=0.0,
                    max_value=round(slider_max_bound, -1),
                    value=(float(min_rate_val), float(max_rate_val)),
                    step=1.0,
                    format="%.0f%%"
                )

                fig_combo = make_subplots(specs=[[{"secondary_y": True}]])
                fig_combo.add_trace(go.Bar(x=combo_data['period'], y=combo_data['유형별_불량수량'], name='총 불량 수량', text=combo_data['유형별_불량수량'], texttemplate='%{text:,.0f}', textposition='auto'), secondary_y=False)
                fig_combo.add_trace(go.Scatter(x=combo_data['period'], y=combo_data['총_불량률(%)'], name='총 불량률 (%)', mode='lines+markers+text', text=combo_data['총_불량률(%)'], texttemplate='%{text:.2f}%', textposition='top center', connectgaps=False, textfont=dict(size=16, color='black')), secondary_y=True)
                fig_combo.update_layout(height=600, title_text=f"<b>{agg_level} 총 불량 수량 및 불량률 추이</b>", legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
                fig_combo.update_yaxes(title_text="<b>총 불량 수량 (개)</b>", secondary_y=False); fig_combo.update_yaxes(title_text="<b>총 불량률 (%)</b>", secondary_y=True, range=rate_range)
                fig_combo.update_xaxes(title_text=f"<b>{agg_level.replace('별', '')}</b>", type='category', categoryorder='array', categoryarray=sorted(combo_data['period'].unique()))
                st.plotly_chart(fig_combo, use_container_width=True)
            else:
                st.info("선택된 필터 조건에 해당하는 추이 데이터가 없습니다.")

            st.divider()
            st.subheader(f"{agg_level} 불량 유형별 불량률 추이", anchor=False)
            
            prod_resampled = get_resampled_data(prod_data_source, agg_level, ['생산수량'], group_by_cols=['period']).rename(columns={'생산수량': '기간별_총생산량'})
            defect_resampled = get_resampled_data(df_display, agg_level, ['유형별_불량수량'], group_by_cols=['period', '불량명'])
            
            if not defect_resampled.empty:
                trend_final_data = pd.merge(defect_resampled, prod_resampled, on='period', how='left')
                production_for_rate_ind = trend_final_data['기간별_총생산량'].replace(0, pd.NA)
                with pd.option_context('mode.use_inf_as_na', True):
                    trend_final_data['불량률(%)'] = (100 * trend_final_data['유형별_불량수량'] / production_for_rate_ind).fillna(0)

                chart_option_cols = st.columns([2, 1, 1])
                with chart_option_cols[0]:
                     top_n_defects = st.number_input(
                         "상위 N개 불량 유형 표시", 
                         min_value=1, 
                         max_value=len(trend_final_data['불량명'].unique()), 
                         value=len(trend_final_data['불량명'].unique()), 
                         step=1,
                         help="평균 불량률이 높은 순으로 상위 N개 유형의 추이만 표시합니다."
                     )
                with chart_option_cols[1]:
                    st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
                    show_labels = st.toggle("차트 라벨 표시", value=True)

                avg_defect_rates = trend_final_data.groupby('불량명')['불량률(%)'].mean().nlargest(top_n_defects).index.tolist()
                trend_final_data_top_n = trend_final_data[trend_final_data['불량명'].isin(avg_defect_rates)]
                
                fig_trend_rate = px.line(trend_final_data_top_n.sort_values('period'), x='period', y='불량률(%)', color='불량명', title=f"<b>{agg_level} 불량 유형별 불량률 추이</b>", markers=True, text='불량률(%)' if show_labels else None, height=600)
                fig_trend_rate.update_traces(texttemplate='%{text:.4f}%', textposition='top center', textfont=dict(size=16, color='black'), connectgaps=False)
                fig_trend_rate.update_layout(legend_title_text='불량 유형', xaxis_title=f"<b>{agg_level.replace('별', '')}</b>", yaxis_title="<b>불량률 (%)</b>")
                fig_trend_rate.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(trend_final_data_top_n['period'].unique()))
                st.plotly_chart(fig_trend_rate, use_container_width=True)
            else:
                st.info("선택된 필터 조건에 해당하는 추이 데이터가 없습니다.")

            with side_col:
                st.markdown(analyze_defect_data(df_defect_filtered))
                st.divider()
                st.subheader("데이터 원본 (필터링됨)")
                st.dataframe(df_display, use_container_width=True, height=500)

elif selected_tab == "가동률 분석":
    df_filtered, _, _, agg_level = create_shared_filter_controls(df_utilization_orig)
    if not df_filtered.empty:
        df_total_agg = aggregate_overall_data(df_filtered, 'utilization'); main_col, side_col = st.columns([2.8, 1]);
        with main_col:
            df_resampled_util = get_resampled_data(df_filtered, agg_level, ['총_생산수량', '이론상_총_생산량'], group_by_cols=['period', '공장', '공정코드'])
            if not df_resampled_util.empty:
                st.subheader(f"{agg_level} 공장별 가동률 추이")
                with pd.option_context('mode.use_inf_as_na', True): df_resampled_util['평균_가동률'] = (100 * df_resampled_util['총_생산수량'] / df_resampled_util['이론상_총_생산량']).fillna(0)
                df_trend = df_resampled_util.groupby(['period', '공장'])['평균_가동률'].mean().reset_index()
                fig_trend = px.line(df_trend.sort_values('period'), x='period', y='평균_가동률', color='공장', title=f'<b>{agg_level} 공장 가동률 추이</b>', markers=True, text='평균_가동률')
                fig_trend.update_traces(texttemplate='%{text:.2f}%', textposition='top center', textfont=dict(size=16, color='black')); fig_trend.update_xaxes(type='category', categoryorder='array', categoryarray=sorted(df_trend['period'].unique())); st.plotly_chart(fig_trend, use_container_width=True)
            all_factories_in_period = sorted(df_filtered['공장'].unique())
            plot_horizontal_bar_chart_all_processes(df_total_agg, {'rate_col': '평균_가동률', 'y_axis_title': '평균 가동률', 'chart_title': '공장/공정별 평균 가동률'}, all_factories_in_period, PROCESS_MASTER_ORDER)
        with side_col: st.markdown(analyze_utilization_data(df_total_agg)); st.divider(); st.subheader("데이터 원본"); st.dataframe(df_filtered, use_container_width=True, height=500)
    else: st.info(f"선택된 기간에 해당하는 가동률 데이터가 없습니다.")

elif selected_tab == "종합 분석":
    df_filtered, start_date, end_date, agg_level = create_shared_filter_controls(df_target_orig)
    if df_filtered.empty or df_yield_orig.empty: st.info("분석에 필요한 목표 달성률 또는 수율 데이터가 없습니다.")
    else:
        mask_yield = (df_yield_orig['date'].dt.date >= start_date) & (df_yield_orig['date'].dt.date <= end_date)
        df_yield_filt = df_yield_orig[mask_yield].copy()

        # 데이터 처리
        compare_factories = st.session_state.get('compare_factories', False)
        selected_factory = st.session_state.get('overall_factory_select', '전체')
        
        if compare_factories:
            df_yield_filt_factory = df_yield_filt.copy()
            active_factory = '전체'
        else:
            df_yield_filt_factory = df_yield_filt[df_yield_filt['공장'] == selected_factory].copy() if selected_factory != '전체' else df_yield_filt.copy()
            active_factory = selected_factory

        bar_data, line_data = pd.DataFrame(), pd.DataFrame()
        if not df_yield_filt_factory.empty:
            group_by_cols = ['period', '공장', '공정코드'] if compare_factories else ['period', '공정코드']
            df_yield_resampled = get_resampled_data(df_yield_filt_factory, agg_level, ['총_생산수량', '총_양품수량'], group_by_cols=group_by_cols)
            df_final_yield_filtered = df_yield_resampled[df_yield_resampled['공정코드'] == '[80] 누수/규격검사']
            bar_group_cols = ['period', '공장'] if compare_factories else ['period']
            bar_data = df_final_yield_filtered.groupby(bar_group_cols)['총_양품수량'].sum().reset_index().rename(columns={'총_양품수량': '총_생산수량'})
            with pd.option_context('mode.use_inf_as_na', True): df_yield_resampled['개별공정수율'] = (df_yield_resampled['총_양품수량'] / df_yield_resampled['총_생산수량']).fillna(1.0)
            line_group_cols = ['period', '공장'] if compare_factories else ['period']
            line_data = df_yield_resampled.groupby(line_group_cols)['개별공정수율'].prod().reset_index(name='종합수율(%)')
            line_data['종합수율(%)'] *= 100
        else:
            bar_data = pd.DataFrame(columns=['period', '총_생산수량'])
            line_data = pd.DataFrame(columns=['period', '종합수율(%)'])

        if bar_data.empty or line_data.empty: st.info("선택된 기간에 분석할 데이터가 부족합니다.")
        else:
            merge_cols = ['period', '공장'] if compare_factories else ['period']
            combo_data = pd.merge(bar_data, line_data, on=merge_cols, how='outer').sort_values('period').fillna(0)
            
            st.markdown("---"); st.subheader("차트 옵션 조정", anchor=False)
            
            # 모든 컨트롤을 브리핑 위로 이동
            control_cols_1 = st.columns(3)
            with control_cols_1[0]:
                all_factories = ['전체'] + sorted(df_yield_orig['공장'].unique())
                st.selectbox(
                    "공장 선택", options=all_factories, key="overall_factory_select",
                    disabled=st.session_state.get('compare_factories', False)
                )
            with control_cols_1[1]:
                st.markdown("<div style='padding-top: 28px;'></div>", unsafe_allow_html=True)
                st.checkbox("공장별 함께보기", key="compare_factories")

            control_cols_2 = st.columns(3)
            with control_cols_2[0]: 
                min_yield_val = combo_data['종합수율(%)'].min() if not combo_data.empty else 0
                max_yield_val = combo_data['종합수율(%)'].max() if not combo_data.empty else 100
                buffer = (max_yield_val - min_yield_val) * 0.5 if max_yield_val > min_yield_val else 5.0
                slider_min = max(0.0, min_yield_val - buffer)
                slider_max = min(100.0, max_yield_val + buffer)
                yield_range = st.slider("종합 수율(%) 축 범위", 0.0, 100.0, (slider_min, slider_max), 1.0, format="%.0f%%", key="overall_yield_range")
            with control_cols_2[1]: chart_height = st.slider("차트 높이 조절", 400, 1000, 600, 50, key="overall_chart_height")
            with control_cols_2[2]: show_labels = st.toggle("차트 라벨 표시", value=True, key="overall_show_labels")
            
            st.markdown(generate_summary_text(combo_data, agg_level, active_factory), unsafe_allow_html=True)
            
            fig = make_subplots(specs=[[{"secondary_y": True}]])
            chart_title_prefix = f"{active_factory} " if active_factory != '전체' else ""
            
            if compare_factories:
                factory_color_map = {'A관': 'blue', 'C관': 'skyblue', 'S관': 'red'}
                for factory_name in sorted(combo_data['공장'].unique()):
                    df_factory = combo_data[combo_data['공장'] == factory_name]
                    
                    factory_color = 'gray'  # 기본값
                    for key, color in factory_color_map.items():
                        if key in factory_name:
                            factory_color = color
                            break
                    
                    fig.add_trace(go.Bar(
                        x=df_factory['period'], y=df_factory['총_생산수량'], name=f'{factory_name} 완제품', 
                        legendgroup=factory_name, marker_color=factory_color,
                        text=df_factory['총_생산수량'], texttemplate='<b>%{text:,.0f}</b>',
                        textposition='outside' if show_labels else 'none',
                        textfont=dict(size=22, color='black')
                    ), secondary_y=False)
                    fig.add_trace(go.Scatter(
                        x=df_factory['period'], y=df_factory['종합수율(%)'], name=f'{factory_name} 수율', 
                        legendgroup=factory_name, line=dict(color=factory_color, dash='dot'), 
                        mode='lines+markers+text' if show_labels else 'lines+markers',
                        text=df_factory['종합수율(%)'], texttemplate='<b>%{text:.2f}%</b>',
                        textposition='top center',
                        textfont=dict(color='black', size=14)
                    ), secondary_y=True)
                fig.update_layout(barmode='group')
            else:
                blue_scale = ['#aed6f1', '#85c1e9', '#5dade2', '#3498db', '#2e86c1', '#2874a6', '#21618c', '#1b4f72', '#153d5a', '#102e48', '#0b1e34', '#071323']
                bar_colors = [blue_scale[i % len(blue_scale)] for i in range(len(combo_data))]
                fig.add_trace(go.Bar(x=combo_data['period'], y=combo_data['총_생산수량'], name='완제품 제조 개수', text=combo_data['총_생산수량'], texttemplate='<b>%{text:,.0f}</b>', textposition='outside' if show_labels else 'none', textfont=dict(size=22), marker_color=bar_colors), secondary_y=False)
                fig.add_trace(go.Scatter(x=combo_data['period'], y=combo_data['종합수율(%)'], name=f'{agg_level} 종합 수율', mode='lines+markers+text' if show_labels else 'lines+markers', line=dict(color='crimson', width=3), marker=dict(color='crimson', size=8), text=combo_data['종합수율(%)'], texttemplate='<b>%{text:.2f}%</b>', textposition='top center', textfont=dict(color='black', size=20, family="Arial, sans-serif")), secondary_y=True)

            max_bar_val = combo_data['총_생산수량'].max() if not combo_data.empty else 0

            fig.update_layout(height=chart_height, title_text=f'<b>{chart_title_prefix}{agg_level} 완제품 제조 실적 및 종합 수율</b>', title_font_size=24, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font_size=16))
            fig.update_yaxes(title_text="<b>완제품 제조 개수</b>", secondary_y=False, title_font_size=18, tickfont_size=14, range=[0, max_bar_val * 1.15])
            fig.update_yaxes(title_text="<b>종합 수율 (%)</b>", secondary_y=True, title_font_size=18, tickfont_size=14, range=yield_range)
            fig.update_xaxes(title_text=f"<b>{agg_level.replace('별', '')}</b>", type='category', categoryorder='array', categoryarray=sorted(combo_data['period'].unique()), title_font_size=18, tickfont_size=14)
            st.plotly_chart(fig, use_container_width=True)


            # --- 제품군별 종합 실적 분석 ---
            st.divider()
            st.subheader(f"{agg_level} 제품군별 완제품 제조 실적 및 종합 수율", anchor=False)

            # 공장 선택 필터
            pg_all_factories = ['전체'] + sorted(df_yield_orig['공장'].unique())
            pg_selected_factory = st.selectbox(
                "분석 공장 선택", 
                options=pg_all_factories, 
                key="pg_factory_select",
                help="제품군별 분석을 수행할 공장을 선택합니다. '전체' 선택 시 모든 공장의 데이터를 종합하여 분석합니다."
            )

            # 선택된 공장에 따라 데이터 필터링
            if pg_selected_factory == '전체':
                df_yield_pg_filtered = df_yield_filt.copy()
            else:
                df_yield_pg_filtered = df_yield_filt[df_yield_filt['공장'] == pg_selected_factory].copy()
            
            if '신규분류요약' in df_yield_pg_filtered.columns:
                all_product_groups_pg = sorted(df_yield_pg_filtered['신규분류요약'].dropna().unique())

                if not all_product_groups_pg:
                    st.warning("선택된 공장에 제품군 데이터가 없습니다.")
                else:
                    for group in all_product_groups_pg:
                        if f"pg_product_group_{group}" not in st.session_state: st.session_state[f"pg_product_group_{group}"] = True
                    
                    st.markdown("##### 표시할 제품군 선택")
                    btn_cols_pg = st.columns(8)
                    with btn_cols_pg[0]:
                        if st.button("제품군 전체 선택", key="pg_select_all", use_container_width=True):
                            for group in all_product_groups_pg: st.session_state[f"pg_product_group_{group}"] = True
                            st.rerun()
                    with btn_cols_pg[1]:
                        if st.button("제품군 전체 해제", key="pg_deselect_all", use_container_width=True):
                            for group in all_product_groups_pg: st.session_state[f"pg_product_group_{group}"] = False
                            st.rerun()
                    
                    st.write("")
                    num_cols_pg = 5
                    cols_pg = st.columns(num_cols_pg)
                    selected_product_groups_pg = []
                    for i, group in enumerate(all_product_groups_pg):
                        with cols_pg[i % num_cols_pg]:
                            if st.checkbox(group, key=f"pg_product_group_{group}"):
                                selected_product_groups_pg.append(group)
                    
                    combine_pg = st.checkbox("선택항목 합쳐서 보기", key="pg_combine_yield", help="선택한 제품군들의 실적을 합산하여 단일 종합 수율 및 생산 실적 추이를 분석합니다.")

                    if selected_product_groups_pg:
                        df_resampled_pg = get_resampled_data(df_yield_pg_filtered, agg_level, ['총_생산수량', '총_양품수량'], group_by_cols=['period', '신규분류요약', '공정코드'])
                        df_resampled_pg_filtered = df_resampled_pg[df_resampled_pg['신규분류요약'].isin(selected_product_groups_pg)]

                        if not df_resampled_pg_filtered.empty:
                            df_to_plot_pg = pd.DataFrame()
                            if combine_pg:
                                bar_combined = df_resampled_pg_filtered[df_resampled_pg_filtered['공정코드'] == '[80] 누수/규격검사'].groupby('period')['총_양품수량'].sum().reset_index().rename(columns={'총_양품수량': '완제품_제조개수'})
                                
                                df_yield_combined_base = df_resampled_pg_filtered.groupby(['period', '공정코드']).agg(총_생산수량=('총_생산수량', 'sum'), 총_양품수량=('총_양품수량', 'sum')).reset_index()
                                with pd.option_context('mode.use_inf_as_na', True): df_yield_combined_base['개별수율'] = (df_yield_combined_base['총_양품수량'] / df_yield_combined_base['총_생산수량']).fillna(1.0)
                                line_combined = df_yield_combined_base.groupby('period')['개별수율'].prod().reset_index(name='종합수율(%)')
                                line_combined['종합수율(%)'] *= 100
                                
                                df_to_plot_pg = pd.merge(bar_combined, line_combined, on='period', how='outer').fillna(0)
                                df_to_plot_pg['신규분류요약'] = "선택항목 종합"
                            else:
                                bar_data_pg = df_resampled_pg_filtered[df_resampled_pg_filtered['공정코드'] == '[80] 누수/규격검사'].groupby(['period', '신규분류요약'])['총_양품수량'].sum().reset_index().rename(columns={'총_양품수량': '완제품_제조개수'})
                                
                                with pd.option_context('mode.use_inf_as_na', True): df_resampled_pg_filtered['개별공정수율'] = (df_resampled_pg_filtered['총_양품수량'] / df_resampled_pg_filtered['총_생산수량']).fillna(1.0)
                                line_data_pg = df_resampled_pg_filtered.groupby(['period', '신규분류요약'])['개별공정수율'].prod().reset_index(name='종합수율(%)')
                                line_data_pg['종합수율(%)'] *= 100
                                
                                df_to_plot_pg = pd.merge(bar_data_pg, line_data_pg, on=['period', '신규분류요약'], how='outer').sort_values('period').fillna(0)

                            if not df_to_plot_pg.empty:
                                fig_pg = make_subplots(specs=[[{"secondary_y": True}]])
                                
                                colors = px.colors.qualitative.Plotly
                                group_col = '신규분류요약'
                                
                                for i, group_name in enumerate(df_to_plot_pg[group_col].unique()):
                                    df_group = df_to_plot_pg[df_to_plot_pg[group_col] == group_name]
                                    color = colors[i % len(colors)]
                                    
                                    fig_pg.add_trace(go.Bar(x=df_group['period'], y=df_group['완제품_제조개수'], name=f'{group_name} 완제품', legendgroup=group_name, marker_color=color), secondary_y=False)
                                    fig_pg.add_trace(go.Scatter(x=df_group['period'], y=df_group['종합수율(%)'], name=f'{group_name} 수율', legendgroup=group_name, mode='lines+markers', line=dict(color=color, dash='dot')), secondary_y=True)

                                factory_title = f"({pg_selected_factory})" if pg_selected_factory != '전체' else '(전체 공장)'
                                fig_pg.update_layout(height=600, title_text=f'<b>{agg_level} 제품군별 완제품 제조 실적 및 종합 수율 {factory_title}</b>', barmode='group', legend_title_text='범례')
                                max_bar_val_pg = df_to_plot_pg['완제품_제조개수'].max() if not df_to_plot_pg.empty else 0
                                fig_pg.update_yaxes(title_text="<b>완제품 제조 개수</b>", secondary_y=False, range=[0, max_bar_val_pg * 1.15]); fig_pg.update_yaxes(title_text="<b>종합 수율 (%)</b>", secondary_y=True, range=[0, 101])
                                fig_pg.update_xaxes(title_text=f"<b>{agg_level.replace('별', '')}</b>", type='category', categoryorder='array', categoryarray=sorted(df_to_plot_pg['period'].unique()))
                                st.plotly_chart(fig_pg, use_container_width=True)
                            else:
                                st.info("선택된 조건에 해당하는 데이터가 없습니다.")
                        else:
                            st.info("선택된 제품군에 대한 데이터가 없습니다.")
                    else:
                        st.info("차트를 표시할 제품군을 선택해주세요.")
            else:
                st.warning("수율 데이터에 '신규분류요약' 컬럼이 없어 제품군별 분석을 제공할 수 없습니다.")


elif selected_tab == "저가동 설비":
    st.header("저가동 설비 분석"); st.info("저가동 설비 데이터는 기간 필터가 적용되지 않고, 로드된 파일의 전체 기간을 기준으로 분석합니다.")
    df_low_util = df_low_util_orig.copy()
    if not df_low_util.empty:
        main_col, side_col = st.columns([2.8, 1])
        with main_col:
            st.subheader("핵심 지표 요약"); kpi1, kpi2, kpi3 = st.columns(3)
            if '기간 내 가동률(%)' in df_low_util.columns and df_low_util['기간 내 가동률(%)'].notna().any():
                worst_performer = df_low_util.loc[df_low_util['기간 내 가동률(%)'].idxmin()]
                kpi2.metric("저가동 설비 평균 가동률", f"{df_low_util['기간 내 가동률(%)'].mean():.2f} %"); kpi3.metric("최저 가동률 설비", f"{worst_performer['기계코드']}", help=f"가동률: {worst_performer['기간 내 가동률(%)']:.2f}%")
            kpi1.metric("총 저가동 설비 수", f"{len(df_low_util)} 대", delta_color="inverse"); st.divider(); st.subheader("저가동 설비 현황")
            df_low_util_sorted = df_low_util.sort_values('기간 내 가동률(%)')
            fig = go.Figure()
            fig.add_trace(go.Bar(x=df_low_util_sorted['기계코드'], y=df_low_util_sorted['기간 내 가동률(%)'], name='실제 가동률', text=df_low_util_sorted['기간 내 가동률(%)'], texttemplate='%{text:.2f}', textposition='outside', marker_color='tomato'))
            fig.add_trace(go.Scatter(x=df_low_util_sorted['기계코드'], y=df_low_util_sorted['저가동설비기준'], name='저가동 기준', mode='lines+markers+text', text=df_low_util_sorted['저가동설비기준'], texttemplate='%{text:.0f}', textposition='top center', line=dict(dash='dot', color='gray'), textfont=dict(size=16, color='black')))
            fig.update_layout(title_text='<b>저가동 설비별 실제 가동률 vs 기준</b>', yaxis_title="가동률 (%)", xaxis_title="설비 코드", height=600); st.plotly_chart(fig, use_container_width=True)
        with side_col: st.markdown(analyze_low_utilization_data(df_low_util)); st.divider(); st.subheader("데이터 원본"); st.dataframe(df_low_util, use_container_width=True, height=500)
    else: st.markdown(analyze_low_utilization_data(df_low_util_orig)); st.success("분석 기간 내 기준 미달인 저가동 설비가 없습니다.")
