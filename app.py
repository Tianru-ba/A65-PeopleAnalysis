import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime
import io

# --- 页面配置 ---
st.set_page_config(
    page_title="A65 项目人力成本分析看板",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# --- V2.0 UI/UX 体验升级版（互联网大厂 B 端视觉规范） ---
TECH_COLORS = {
    'primary': '#0D6EFD',
    'secondary': '#92CC76',
    'warning': '#FFC552',
    'danger': '#E86452',
    'neutral': '#6C757D',
    'background': '#FFFFFF',
    'card_bg': '#FFFFFF',
    'text_main': '#212529',
    'text_sub': '#6C757D',
    'border': 'rgba(0,0,0,0.06)',
    'shadow': '0 12px 30px rgba(0,0,0,0.06)',
    'active_shadow': '0 0 0 4px rgba(13,110,253,0.18)'
}

# --- 样式定制 (V2.0：高呼吸感 + 严谨字体层级 + 渐进式信息暴露) ---
st.markdown(f"""
    <style>
    html, body, [data-testid="stAppViewContainer"] {{
        font-size: 20px !important;
        background-color: {TECH_COLORS['background']};
        color: {TECH_COLORS['text_main']};
    }}

    [data-testid="stAppViewContainer"] .main .block-container {{
        max-width: 1280px;
        padding-left: 1.25rem;
        padding-right: 1.25rem;
    }}

    header[data-testid="stHeader"] {{
        background: rgba(255,255,255,0.95) !important;
        backdrop-filter: blur(8px);
        -webkit-backdrop-filter: blur(8px);
        border-bottom: 1px solid {TECH_COLORS['border']};
    }}

    /* Typography */
    html, body, [data-testid="stAppViewContainer"] {{
        background-color: {TECH_COLORS['background']};
        color: {TECH_COLORS['text_main']};
    }}
    
    h1 {{ 
        font-size: 36px !important;
        font-weight: 700 !important;
        letter-spacing: 0.2px !important;
        color: {TECH_COLORS['text_main']} !important;
        margin-bottom: 0.25rem !important;
    }}
    h2 {{ 
        font-size: 30px !important;
        font-weight: 800 !important;
        color: {TECH_COLORS['text_main']} !important;
        margin-top: 1.25rem !important;
        margin-bottom: 0.75rem !important;
    }}
    h3 {{ 
        font-size: 22px !important;
        font-weight: 800 !important;
        color: {TECH_COLORS['text_main']} !important;
        margin-top: 1rem !important;
    }}

    .meta-line {{
        font-size: 20px;
        color: {TECH_COLORS['neutral']};
        margin-top: 0.25rem;
        margin-bottom: 1rem;
    }}
    
    /* Metric Cards */
    div[data-testid="stMetric"] {{
        background-color: {TECH_COLORS['card_bg']} !important;
        padding: 18px 20px !important;
        border-radius: 12px !important;
        box-shadow: {TECH_COLORS['shadow']} !important;
        border: none !important;
        transition: transform 140ms ease, box-shadow 140ms ease;
    }}
    div[data-testid="stMetric"]:hover {{
        transform: translateY(-2px);
        box-shadow: 0 16px 44px rgba(0,0,0,0.10), 0 0 0 4px rgba(13,110,253,0.10) !important;
    }}
    div[data-testid="stMetric"]:active {{
        transform: translateY(-1px);
    }}
    div[data-testid="stMetricLabel"] > div {{
        font-size: 24px !important;
        font-weight: 900 !important;
        color: {TECH_COLORS['text_main']} !important;
    }}
    div[data-testid="stMetricValue"] > div {{
        font-size: 42px !important;
        font-weight: 800 !important;
        color: {TECH_COLORS['primary']} !important;
    }}

    div[data-testid="stHorizontalBlock"] {{
        gap: 14px;
    }}
    div[data-testid="column"] {{
        padding-left: 6px;
        padding-right: 6px;
    }}

    hr {{
        border: none;
        border-top: 1px solid {TECH_COLORS['border']};
        margin: 18px 0;
    }}
    
    /* Sidebar (Drawer-like) */
    section[data-testid="stSidebar"] {{
        border-right: 1px solid {TECH_COLORS['border']};
        background: #FFFFFF !important;
    }}
    section[data-testid="stSidebar"] {{
        font-size: 16px !important;
    }}
    section[data-testid="stSidebar"] h1 {{
        font-size: 24px !important;
        font-weight: 800 !important;
    }}
    section[data-testid="stSidebar"] h2 {{
        font-size: 18px !important;
    }}
    .stMultiSelect label, .stSelectbox label {{
        font-size: 16px !important;
        font-weight: 700 !important;
    }}
    
    /* Tabs -> Nav Pills */
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {{
        gap: 8px;
        padding: 6px;
        border-radius: 999px;
        background: rgba(255,255,255,0.85);
        border: 1px solid {TECH_COLORS['border']};
        box-shadow: {TECH_COLORS['shadow']};
        border-bottom: none !important;
    }}
    /* 去掉 Tabs 默认的底部指示线（截图中的红色底线） */
    div[data-testid="stTabs"] [data-baseweb="tab-highlight"] {{
        display: none !important;
    }}
    div[data-testid="stTabs"] [data-baseweb="tab-border"] {{
        display: none !important;
    }}
    div[data-testid="stTabs"] [data-baseweb="tab"] {{
        padding: 8px 14px !important;
        border-radius: 999px !important;
        font-size: 20px !important;
        font-weight: 800 !important;
        color: {TECH_COLORS['text_sub']} !important;
        background: transparent !important;
        border: none !important;
        transition: background-color 120ms ease, box-shadow 120ms ease, transform 120ms ease;
    }}
    div[data-testid="stTabs"] [data-baseweb="tab"]:hover {{
        background: rgba(13,110,253,0.08) !important;
    }}
    div[data-testid="stTabs"] [aria-selected="true"] {{
        color: #FFFFFF !important;
        background: {TECH_COLORS['primary']} !important;
        box-shadow: 0 6px 16px rgba(13,110,253,0.18) !important;
        transform: translateY(-1px);
    }}

    section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
        background: #FFFFFF !important;
        border: 1px solid {TECH_COLORS['border']} !important;
        border-radius: 10px !important;
        box-shadow: none !important;
    }}
    section[data-testid="stSidebar"] [data-baseweb="select"] {{
        font-size: 16px !important;
    }}
    section[data-testid="stSidebar"] [data-baseweb="select"] input {{
        font-size: 16px !important;
    }}
    section[data-testid="stSidebar"] [data-baseweb="select"] > div:focus-within {{
        border-color: rgba(13,110,253,0.45) !important;
        box-shadow: {TECH_COLORS['active_shadow']} !important;
    }}
    section[data-testid="stSidebar"] [data-baseweb="tag"] {{
        background: rgba(13,110,253,0.08) !important;
        border: 1px solid rgba(13,110,253,0.18) !important;
        color: {TECH_COLORS['primary']} !important;
        border-radius: 999px !important;
        font-weight: 700 !important;
        font-size: 14px !important;
    }}
    section[data-testid="stSidebar"] [data-baseweb="tag"] span {{
        color: {TECH_COLORS['primary']} !important;
    }}
    section[data-testid="stSidebar"] [role="option"][aria-selected="true"] {{
        background: rgba(13,110,253,0.10) !important;
        color: {TECH_COLORS['text_main']} !important;
    }}
    section[data-testid="stSidebar"] [role="option"]:hover {{
        background: rgba(13,110,253,0.06) !important;
    }}

    /* Guide block */
    .guide-box {{
        background-color: #F0F5FF;
        border: 1px solid {TECH_COLORS['border']};
        border-radius: 10px;
        padding: 8px 12px;
        margin: 6px 0 14px 0;
    }}
    .guide-box p {{
        font-size: 16px;
        font-weight: 500;
        color: {TECH_COLORS['text_sub']};
        margin: 0;
    }}

    div[data-testid="stPlotlyChart"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: {TECH_COLORS['shadow']} !important;
        border-radius: 14px;
        transition: transform 140ms ease, box-shadow 140ms ease;
    }}
    div[data-testid="stPlotlyChart"]:hover {{
        transform: translateY(-1px);
        box-shadow: 0 16px 44px rgba(0,0,0,0.10) !important;
    }}

    details[data-testid="stExpander"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        transition: none;
    }}

    div[data-testid="stDataFrame"] {{
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        transition: none;
    }}
    </style>
    """, unsafe_allow_html=True)

# --- 数据加载与处理 ---
@st.cache_data
def load_data():
    file_path = "A65汇总及明细.xlsx"
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names
    
    # 智能匹配 Sheet 名称
    detail_sheet = next((s for s in sheet_names if '明细' in s and '预测' not in s), None)
    if not detail_sheet:
        detail_sheet = next((s for s in sheet_names if 'Sheet1' in s), sheet_names[0])
    
    summary_sheet = next((s for s in sheet_names if '汇总' in s), None)
    forecast_sheet = next((s for s in sheet_names if '预测' in s), None)
    
    # 读取明细表
    df_detail = pd.read_excel(file_path, sheet_name=detail_sheet)
    
    # --- 强力列名自动适配逻辑 ---
    col_map = {
        '月份': ['月份', '统计年月', '日期', 'month', 'date', '时间', 'time'],
        '姓名': ['姓名', 'name', '人员', 'person'],
        '职能分类': ['职能分类', '岗位分类', '职能', 'category', '岗位', 'class'],
        '每月总人力(USD)': ['每月总人力(USD)', '总成本', '总人力', 'total cost', 'usd', '人力成本', 'cost'],
        '固定薪资(USD)': ['固定薪资(USD)', '固定薪资', 'base salary', '固定', 'base'],
        '奖金/激励(USD)': ['奖金/激励(USD)', '奖金', '激励', 'incentive', 'bonus'],
        '即时激励(USD)': ['即时激励(USD)', '即时激励', 'spot bonus', 'spot'],
        '岗位性质': ['岗位性质', '用工形式', '全职/兼职', 'Type', 'Status']
    }
    
    def auto_rename_cols(df, mapping):
        actual_cols = df.columns.tolist()
        rename_dict = {}
        used_actual_cols = set()
        for standard_name, possible_names in mapping.items():
            match = next((c for c in actual_cols if str(c).strip() == standard_name), None)
            if not match:
                match = next((c for c in actual_cols if c not in used_actual_cols and any(p.lower() in str(c).lower() for p in possible_names)), None)
            if match and match not in used_actual_cols:
                rename_dict[match] = standard_name
                used_actual_cols.add(match)
        return df.rename(columns=rename_dict).loc[:, ~df.rename(columns=rename_dict).columns.duplicated()]

    df_detail = auto_rename_cols(df_detail, col_map)
    
    # --- 核心数据清洗 ---
    def robust_parse_date(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x, (int, float)):
            try: return pd.to_datetime(x, unit='D', origin='1899-12-30')
            except: pass
        try: return pd.to_datetime(x)
        except:
            try: return pd.to_datetime(str(x).replace('.', '-'))
            except: return pd.NaT

    df_detail['月份'] = df_detail['月份'].apply(robust_parse_date)
    df_detail = df_detail.dropna(subset=['月份']).sort_values('月份')
    
    # 衍生时间维度
    df_detail['年份'] = df_detail['月份'].dt.year
    df_detail['季度'] = df_detail['月份'].dt.quarter
    df_detail['月'] = df_detail['月份'].dt.month
    
    # 数值转换
    cost_cols = ['每月总人力(USD)', '固定薪资(USD)', '奖金/激励(USD)', '即时激励(USD)']
    for col in cost_cols:
        if col in df_detail.columns:
            df_detail[col] = pd.to_numeric(df_detail[col], errors='coerce').fillna(0)
    
    # 职能映射
    mapping = {'CEO': '内容创作', '产品经理': '内容创作', 'BD&CM': '制作执行', '综合岗': '支持运营'}
    df_detail['职能大类'] = df_detail['职能分类'].map(mapping).fillna('其他')
    
    # --- 处理预测数据 (来自新 Sheet "预测明细") ---
    df_summary_data = pd.DataFrame()
    if forecast_sheet:
        df_forecast = pd.read_excel(file_path, sheet_name=forecast_sheet)
        # 尝试适配预测表的列名
        f_col_map = {
            '月份': ['月份', '统计年月', '日期', '时间'],
            '固定薪酬': ['固定薪酬', '固定薪资', '固定'],
            '奖金': ['奖金', '奖金/激励', '激励'],
            '即时激励': ['即时激励'],
            '人数': ['人数', 'A65直接人数合计', '人员合计']
        }
        df_forecast = auto_rename_cols(df_forecast, f_col_map)
        if '月份' in df_forecast.columns:
            df_forecast['月份'] = df_forecast['月份'].apply(robust_parse_date)
            df_forecast = df_forecast.dropna(subset=['月份'])
            df_forecast['数据类型'] = '预测'
            # 数值转换
            for c in ['固定薪酬', '奖金', '即时激励', '人数']:
                if c in df_forecast.columns:
                    df_forecast[c] = pd.to_numeric(df_forecast[c], errors='coerce').fillna(0)
            df_summary_data = df_forecast
    
    return df_detail, df_summary_data

try:
    df_detail, df_summary = load_data()
except Exception as e:
    st.error(f"数据处理过程出错: {e}")
    st.stop()

# --- 侧边栏过滤器 (V2.0：直接展示) ---
st.sidebar.title("筛选器")
st.sidebar.subheader("时间维度")
years = sorted(df_detail['年份'].unique().tolist())
sel_years = st.sidebar.multiselect("年度", options=years, default=years)

if sel_years:
    df_temp_y = df_detail[df_detail['年份'].isin(sel_years)]
    available_quarters = sorted(df_temp_y['季度'].unique().tolist())
else:
    available_quarters = sorted(df_detail['季度'].unique().tolist())
sel_quarters = st.sidebar.multiselect(
    "季度",
    options=[f"Q{q}" for q in available_quarters],
    default=[f"Q{q}" for q in available_quarters],
)

mask_m = pd.Series([True] * len(df_detail))
if sel_years:
    mask_m &= (df_detail['年份'].isin(sel_years))
if sel_quarters:
    q_nums = [int(q[1]) for q in sel_quarters]
    mask_m &= (df_detail['季度'].isin(q_nums))
available_months = sorted(df_detail[mask_m]['月'].unique().tolist())
sel_months = st.sidebar.multiselect("月份", options=available_months, default=available_months)

st.sidebar.subheader("职能分类")
all_categories = sorted(df_detail['职能分类'].unique().tolist())
selected_categories = st.sidebar.multiselect("职能", options=all_categories, default=all_categories)

# --- 执行最终过滤 ---
df_filtered = df_detail.copy()
if sel_years:
    df_filtered = df_filtered[df_filtered['年份'].isin(sel_years)]
if sel_quarters:
    q_nums = [int(q[1]) for q in sel_quarters]
    df_filtered = df_filtered[df_filtered['季度'].isin(q_nums)]
if sel_months:
    df_filtered = df_filtered[df_filtered['月'].isin(sel_months)]
df_filtered = df_filtered[df_filtered['职能分类'].isin(selected_categories)]

# --- 页面内容 ---
st.title("A65 人力成本分析仪表盘")
st.markdown(
    f"<div class='meta-line'>数据源：A65汇总及明细.xlsx ｜统计周期：{df_filtered['月份'].min():%Y.%m} - {df_filtered['月份'].max():%Y.%m} ｜记录数：{len(df_filtered)}</div>",
    unsafe_allow_html=True
)

# 顶部指标卡 (增加对比逻辑)
if not df_filtered.empty:
    latest_date = df_filtered['月份'].max()
    curr_data = df_filtered[df_filtered['月份'] == latest_date]
    
    total_all_cost = df_detail['每月总人力(USD)'].sum() if '每月总人力(USD)' in df_detail.columns else 0
    total_hc = len(curr_data)
    total_last_month_cost = curr_data['每月总人力(USD)'].sum()
    total_period_cost = df_filtered['每月总人力(USD)'].sum()
    months_in_period = df_filtered['月份'].dt.to_period('M').nunique()
    avg_period_monthly_cost = (total_period_cost / months_in_period) if months_in_period > 0 else 0
    avg_cost = (total_last_month_cost / total_hc) if total_hc > 0 else 0
    
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    with r1c1:
        st.metric(
            "累计总成本",
            f"${total_all_cost/1000:.1f}k",
            help="计算逻辑：不随筛选条件变化。基于数据源“明细/Sheet1”全量数据的“每月总人力(USD)”累计总和，用于作为项目总盘的固定参照。"
        )
    with r1c2:
        st.metric(
            "分析周期月均成本",
            f"${avg_period_monthly_cost/1000:.1f}k",
            help="计算逻辑：分析周期总成本 ÷ 分析周期月份数（按 YYYY-MM 去重）。用于衡量周期内平均月度消耗水平。"
        )
    with r1c3:
        st.metric(
            "分析周期末总人数",
            f"{total_hc} 人",
            help="计算逻辑：当前所选时间范围中最后一个月的总人数（Headcount）。"
        )
    with r1c4:
        st.metric(
            "分析周期末总成本",
            f"${total_last_month_cost/1000:.1f}k",
            help="计算逻辑：对应数据源“明细”表中最后一个月的“每月总人力(USD)”列之和。所有数值以 USD 结算。"
        )

    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    with r2c1:
        st.metric(
            "周期末人均单价",
            f"${avg_cost/1000:.1f}k",
            help="计算逻辑：分析周期末总成本 ÷ 分析周期末总人数。"
        )
    with r2c2:
        st.metric(
            "人效",
            "待接入",
            help="预留指标。常见口径：收入/人（或产出/人）。需接入“收入/产出”数据字段后启用。"
        )
    with r2c3:
        st.metric(
            "人工成本率",
            "待接入",
            help="预留指标。常见口径：人工成本 ÷ 收入（或总成本）。需接入“收入”字段并确认口径后启用。"
        )
    with r2c4:
        st.metric(
            "人工成本毛利率",
            "待接入",
            help="预留指标。常见口径：(毛利 - 人工成本) ÷ 毛利 或 1 - 人工成本/毛利。需接入“毛利”字段并确认口径后启用。"
        )
else:
    st.warning("当前筛选条件下无数据，请调整筛选器。")
    st.stop()

st.markdown("---")

# --- 核心图表优化 ---
# --- 核心图表优化 ---
# 统一 Plotly 样式 (专业 UX 交互版：极致突出 + 提亮效果)
def update_fig_layout(fig, title, is_cost_chart=False):
    fig.update_layout(
        title=dict(
            text=f"<b>{title}</b>", 
            font=dict(size=22, color=TECH_COLORS['text_main']), 
            x=0, y=0.98
        ),
        template="plotly_white",
        hovermode="closest", 
        clickmode="event+select",
        font=dict(size=18, color=TECH_COLORS['text_sub']), 
        hoverlabel=dict(
            bgcolor="white",
            font_size=19, 
            font_family="PingFang SC, Microsoft YaHei",
            bordercolor=TECH_COLORS['primary'],
            namelength=-1
        ),
        legend=dict(
            orientation="h", 
            yanchor="bottom", y=1.02, 
            xanchor="right", x=1,
            font=dict(size=18) 
        ),
        margin=dict(l=60, r=60, t=100, b=60),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(
            showgrid=False, 
            tickfont=dict(size=18),
            title_font=dict(size=18),
            linecolor='#E5E6EB',
            showspikes=True,
            spikemode='across',
            spikesnap='cursor',
            showline=True,
            spikedash='dash',
            spikecolor='#999999',
            spikethickness=1
        ),
        yaxis=dict(
            gridcolor='#F0F2F5', 
            tickfont=dict(size=18),
            title_font=dict(size=18),
            zeroline=False,
            showspikes=True,
            spikemode='across',
            spikesnap='cursor',
            showline=True,
            spikedash='dash',
            spikecolor='#999999',
            spikethickness=1,
            tickformat=".1s" if is_cost_chart else None, # 如果是金额图，统一用 k 格式
            tickprefix="$" if is_cost_chart else None
        )
    )
    # 极致交互：设置选中和未选中状态的样式对比
    fig.update_traces(
        # 基础状态：清晰边框
        marker_line_width=2, 
        marker_line_color="white",
        # 交互效果：选中的柱子 100% 提亮，未选中的 20% 透明度（变暗）
        # 注意：bar 的 selected.marker 不支持 line 属性，仅支持 opacity 和 color
        selected=dict(marker=dict(opacity=1.0)),
        unselected=dict(marker=dict(opacity=0.2)),
        selector=dict(type='bar')
    )
    fig.update_traces(
        # 针对折线图的选中突出
        selected=dict(marker=dict(size=20, opacity=1.0)),
        unselected=dict(marker=dict(opacity=0.3)),
        selector=dict(type='scatter')
    )

tab1, tab2, tab3 = st.tabs(["维度一：整体成本", "维度二：职能透视", "维度三：预算执行"])

with tab1:
    st.subheader("月度成本变动监控")
    st.markdown(
        "<div class=\"guide-box\"><p>💡 <b>看数指南</b>：下方柱状图展示本月比上月多花（或少花）了多少钱。红色代表支出增加，绿色代表支出减少。标签显示变动金额及环比增长率。</p></div>",
        unsafe_allow_html=True,
    )

    grouped_cost = df_detail.groupby('月份').agg({'每月总人力(USD)': 'sum'})
    mom_abs_df = grouped_cost.diff().reset_index().rename(columns={'每月总人力(USD)': '成本变动额'})
    mom_pct_df = grouped_cost.pct_change().reset_index().rename(columns={'每月总人力(USD)': '环比率'})
    mom_combined = pd.merge(mom_abs_df, mom_pct_df, on='月份').dropna()

    if not df_filtered.empty:
        f_min_date, f_max_date = df_filtered['月份'].min(), df_filtered['月份'].max()
        mom_display = mom_combined[(mom_combined['月份'] >= f_min_date) & (mom_combined['月份'] <= f_max_date)]
    else:
        mom_display = mom_combined

    fig_mom = go.Figure()
    fig_mom.add_trace(go.Bar(
        x=mom_display['月份'],
        y=mom_display['成本变动额'],
        name='月度变动额',
        marker_color=[TECH_COLORS['danger'] if val >= 0 else TECH_COLORS['secondary'] for val in mom_display['成本变动额']],
        text=mom_display.apply(lambda x: f"<b>{'+' if x['成本变动额'] >= 0 else '-'}${abs(x['成本变动额']/1000):.1f}k</b><br>({x['环比率']:.1%})", axis=1),
        textposition='outside',
        textfont=dict(size=12, color=TECH_COLORS['text_main']),
        cliponaxis=False,
        hovertemplate="月份: %{x|%Y-%m}<br>变动金额: $%{y:,.2f}<extra></extra>"
    ))
    update_fig_layout(fig_mom, "", is_cost_chart=True)
    fig_mom.update_layout(yaxis=dict(dtick=20000), xaxis=dict(tickformat="%Y-%m"), showlegend=False, margin=dict(t=40))
    st.plotly_chart(fig_mom, width="stretch")

    st.subheader("月度成本变动归因拆解")
    st.markdown(
        "<div class=\"guide-box\"><p>💡 <b>看数指南</b>：将每月总变动额拆解为固定薪资、奖金、即时激励，并用主趋势线展示人数环比增减，用于判断“调薪驱动”还是“招人驱动”。</p></div>",
        unsafe_allow_html=True,
    )

    category_diff = df_detail.groupby('月份').agg({
        '固定薪资(USD)': 'sum',
        '奖金/激励(USD)': 'sum',
        '即时激励(USD)': 'sum',
        '姓名': 'count'
    }).diff().reset_index().rename(columns={'姓名': '人数变动'}).dropna()

    if not df_filtered.empty:
        diff_display = category_diff[(category_diff['月份'] >= f_min_date) & (category_diff['月份'] <= f_max_date)]
    else:
        diff_display = category_diff

    fig_diff_structure = make_subplots(specs=[[{"secondary_y": True}]])
    fig_diff_structure.add_trace(go.Bar(
        x=diff_display['月份'], y=diff_display['固定薪资(USD)'], name='固定薪资变动', marker_color=TECH_COLORS['primary'],
        text=diff_display['固定薪资(USD)'].apply(lambda x: f"${x/1000:.1f}k" if abs(x) > 100 else ""), textposition='inside'
    ), secondary_y=False)
    fig_diff_structure.add_trace(go.Bar(
        x=diff_display['月份'], y=diff_display['奖金/激励(USD)'], name='奖金/激励变动', marker_color=TECH_COLORS['secondary'],
        text=diff_display['奖金/激励(USD)'].apply(lambda x: f"${x/1000:.1f}k" if abs(x) > 100 else ""), textposition='inside'
    ), secondary_y=False)
    fig_diff_structure.add_trace(go.Bar(
        x=diff_display['月份'], y=diff_display['即时激励(USD)'], name='即时激励变动', marker_color=TECH_COLORS['warning'],
        text=diff_display['即时激励(USD)'].apply(lambda x: f"${x/1000:.1f}k" if abs(x) > 100 else ""), textposition='inside'
    ), secondary_y=False)
    fig_diff_structure.add_trace(go.Scatter(
        x=diff_display['月份'],
        y=diff_display['人数变动'],
        name='人数变动 (人)',
        line=dict(color=TECH_COLORS['primary'], width=3, shape='spline'),
        marker=dict(size=10, symbol='circle', color='white', line=dict(width=2, color=TECH_COLORS['primary'])),
        mode='lines+markers+text',
        text=diff_display['人数变动'].apply(lambda x: f"+{int(x)}人" if x > 0 else f"{int(x)}人"),
        textposition='top center',
        textfont=dict(size=12, color=TECH_COLORS['primary'])
    ), secondary_y=True)
    update_fig_layout(fig_diff_structure, "", is_cost_chart=True)
    fig_diff_structure.update_layout(barmode='relative', yaxis_title="成本变动金额 (USD)", yaxis2_title="人数变动 (人)", xaxis=dict(tickformat="%Y-%m"))
    st.plotly_chart(fig_diff_structure, width="stretch")

    st.subheader("维度 A：人力规模与薪酬结构趋势")
    st.markdown(
        "<div class=\"guide-box\"><p>💡 <b>看数指南</b>：堆积柱状图展示总成本构成；主趋势线展示总人数。预测段使用浅色/纹理与虚线区分。</p></div>",
        unsafe_allow_html=True,
    )

    actual_trend = df_filtered.groupby('月份').agg({
        '固定薪资(USD)': 'sum',
        '奖金/激励(USD)': 'sum',
        '即时激励(USD)': 'sum',
        '姓名': 'count'
    }).reset_index().rename(columns={'固定薪资(USD)': '固定', '奖金/激励(USD)': '奖金', '即时激励(USD)': '即时激励', '姓名': '人数'})
    actual_trend['类型'] = '实际'
    actual_trend['总成本'] = actual_trend['固定'] + actual_trend['奖金'] + actual_trend['即时激励']

    forecast_trend = pd.DataFrame()
    if not df_summary.empty and '月份' in df_summary.columns:
        f_mask = (df_summary['月份'] >= '2026-04-01')
        required_f_cols = ['固定薪酬', '奖金', '即时激励', '人数']
        existing_f_cols = [c for c in required_f_cols if c in df_summary.columns]
        if existing_f_cols:
            forecast_trend = df_summary[f_mask].copy()
            if not forecast_trend.empty:
                forecast_trend = forecast_trend.rename(columns={'固定薪酬': '固定', '奖金': '奖金', '即时激励': '即时激励'})
                forecast_trend['类型'] = '预测'
                for c in ['人数', '固定', '奖金', '即时激励']:
                    if c in forecast_trend.columns:
                        forecast_trend[c] = pd.to_numeric(forecast_trend[c], errors='coerce').fillna(0)
                forecast_trend['总成本'] = forecast_trend['固定'] + forecast_trend['奖金'] + forecast_trend['即时激励']

    fig_combined = make_subplots(specs=[[{"secondary_y": True}]])
    fig_combined.add_trace(go.Bar(
        x=actual_trend['月份'], y=actual_trend['固定'], name='固定薪资 (实际)', marker_color=TECH_COLORS['primary'],
        text=actual_trend['固定'].apply(lambda x: f"${x/1000:.1f}k"), textposition='inside'
    ), secondary_y=False)
    fig_combined.add_trace(go.Bar(
        x=actual_trend['月份'], y=actual_trend['奖金'], name='奖金/激励 (实际)', marker_color=TECH_COLORS['secondary'],
        text=actual_trend['奖金'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
    ), secondary_y=False)
    fig_combined.add_trace(go.Bar(
        x=actual_trend['月份'], y=actual_trend['即时激励'], name='即时激励 (实际)', marker_color=TECH_COLORS['warning'],
        text=actual_trend['即时激励'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
    ), secondary_y=False)

    if not forecast_trend.empty:
        fig_combined.add_trace(go.Bar(
            x=forecast_trend['月份'], y=forecast_trend['固定'], name='固定薪资 (预测)', marker_color=TECH_COLORS['primary'],
            opacity=0.35, marker_pattern_shape="/", showlegend=False,
            text=forecast_trend['固定'].apply(lambda x: f"${x/1000:.1f}k"), textposition='inside'
        ), secondary_y=False)
        fig_combined.add_trace(go.Bar(
            x=forecast_trend['月份'], y=forecast_trend['奖金'], name='奖金/激励 (预测)', marker_color=TECH_COLORS['secondary'],
            opacity=0.35, marker_pattern_shape="/", showlegend=False,
            text=forecast_trend['奖金'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
        ), secondary_y=False)
        fig_combined.add_trace(go.Bar(
            x=forecast_trend['月份'], y=forecast_trend['即时激励'], name='即时激励 (预测)', marker_color=TECH_COLORS['warning'],
            opacity=0.35, marker_pattern_shape="/", showlegend=False,
            text=forecast_trend['即时激励'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
        ), secondary_y=False)

    fig_combined.add_trace(go.Scatter(
        x=actual_trend['月份'], y=actual_trend['人数'], name="总人数 (实际)",
        line=dict(color=TECH_COLORS['primary'], width=3, shape='spline'),
        marker=dict(size=10, symbol='circle', color='white', line=dict(width=2, color=TECH_COLORS['primary'])),
        mode='lines+markers+text',
        text=actual_trend['人数'].apply(lambda x: f"{int(x)}人"),
        textposition='top center',
        textfont=dict(size=12, color=TECH_COLORS['primary'])
    ), secondary_y=True)

    if not forecast_trend.empty:
        last_actual = actual_trend.tail(1)
        line_forecast = pd.concat([last_actual, forecast_trend]).sort_values('月份')
        fig_combined.add_trace(go.Scatter(
            x=line_forecast['月份'], y=line_forecast['人数'], name="总人数 (预测)",
            line=dict(color=TECH_COLORS['primary'], width=3, dash='dash', shape='spline'),
            marker=dict(size=10, symbol='circle', color='white', line=dict(width=2, color=TECH_COLORS['primary'])),
            mode='lines+markers',
            showlegend=False
        ), secondary_y=True)

    update_fig_layout(fig_combined, "", is_cost_chart=True)
    fig_combined.update_layout(
        barmode='stack',
        yaxis_title="人力成本 (USD)",
        yaxis2_title="人数 (人)",
        xaxis=dict(tickformat="%Y-%m", dtick="M1", tickangle=-45)
    )
    st.plotly_chart(fig_combined, width="stretch")

with tab2:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("维度 B：职能与岗位结构深度穿透")
        if df_filtered.empty:
            st.warning("无数据可展示")
        else:
            latest_month_in_filter = df_filtered['月份'].max()
            sun_data_latest = df_filtered[df_filtered['月份'] == latest_month_in_filter]
            st.markdown(
                f"<div class=\"guide-box\"><p>💡 <b>看数指南</b>：展示 <b>{latest_month_in_filter.strftime('%Y-%m')}</b> 的人员结构快照。内环为职能，外环为岗位性质。点击内环可钻取。</p></div>",
                unsafe_allow_html=True,
            )
            sun_plot_df = sun_data_latest.groupby(['职能分类', '岗位性质']).size().reset_index(name='人数')
            fig_b = px.sunburst(
                sun_plot_df,
                path=['职能分类', '岗位性质'],
                values='人数',
                color='职能分类',
                color_discrete_sequence=px.colors.qualitative.Pastel,
                template="plotly_white"
            )
            fig_b.update_traces(
                textinfo="label+value",
                hovertemplate="<b>%{label}</b><br>人数: %{value}人<br>占比: %{percentRoot:.1%}<extra></extra>"
            )
            fig_b.update_layout(
                margin=dict(t=10, l=10, r=10, b=10),
                font=dict(size=12, color=TECH_COLORS['text_sub']),
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)'
            )
            st.plotly_chart(fig_b, width="stretch")

    with c2:
        st.subheader("维度 C：职能人均成本构成")
        unit_breakdown = df_filtered.groupby('职能分类').agg({
            '固定薪资(USD)': 'sum',
            '奖金/激励(USD)': 'sum',
            '即时激励(USD)': 'sum',
            '姓名': 'count'
        }).reset_index()
        unit_breakdown['人均固定'] = unit_breakdown['固定薪资(USD)'] / unit_breakdown['姓名']
        unit_breakdown['人均奖金'] = unit_breakdown['奖金/激励(USD)'] / unit_breakdown['姓名']
        unit_breakdown['人均即时'] = unit_breakdown['即时激励(USD)'] / unit_breakdown['姓名']
        unit_breakdown['总人均'] = unit_breakdown['人均固定'] + unit_breakdown['人均奖金'] + unit_breakdown['人均即时']
        unit_breakdown = unit_breakdown.sort_values('总人均', ascending=False)

        with st.expander("查看维度 C 计算规则", expanded=False):
            st.markdown(
                "<div class=\"guide-box\"><p>💡 <b>看数指南</b>：展示各职能平均每人每月成本构成，帮助判断人均高是底薪高还是奖金高。</p></div>",
                unsafe_allow_html=True,
            )
            st.write("""
            **计算逻辑：方案 C (职能维度汇总)**
            1. 分项汇总：固定薪资、奖金/激励、即时激励分别求和。
            2. 人数汇总：统计该职能下的总人次（每月人数加总）。
            3. 人均计算：分项总额 / 总人次。
            """)
            check_df = unit_breakdown[['职能分类', '姓名', '固定薪资(USD)', '奖金/激励(USD)', '总人均']].copy()
            check_df.columns = ['职能分类', '累计总人次(月/人)', '固定薪资累计', '奖金/激励累计', '计算出的人均月成本']
            st.dataframe(
                check_df.style.format({
                    '固定薪资累计': '${:,.0f}',
                    '奖金/激励累计': '${:,.0f}',
                    '计算出的人均月成本': '${:,.0f}'
                }),
                use_container_width=True
            )

        fig_c_stack = go.Figure()
        fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均固定'], name='人均固定薪资', marker_color=TECH_COLORS['primary']))
        fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均奖金'], name='人均奖金/激励', marker_color=TECH_COLORS['secondary']))
        fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均即时'], name='人均即时激励', marker_color=TECH_COLORS['warning']))
        fig_c_stack.add_trace(go.Scatter(
            x=unit_breakdown['职能分类'],
            y=unit_breakdown['总人均'],
            mode='text',
            text=unit_breakdown['总人均'].apply(lambda x: f"<b>${x/1000:.1f}k</b>"),
            textposition='top center',
            textfont=dict(size=12, color=TECH_COLORS['text_main']),
            showlegend=False,
            hoverinfo='skip'
        ))
        update_fig_layout(fig_c_stack, "", is_cost_chart=True)
        fig_c_stack.update_layout(
            barmode='stack',
            yaxis_title="人均成本 (USD/月)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
        )
        st.plotly_chart(fig_c_stack, width="stretch")

with tab3:
    st.subheader("维度 D：预测 vs. 实际 (Burn Rate)")
    st.markdown(
        "<div class=\"guide-box\"><p>💡 <b>看数指南</b>：对比每月总成本的实际消耗与未来预测。基准线为历史月均 Burn Rate。</p></div>",
        unsafe_allow_html=True,
    )

    actual_burn = df_detail.groupby('月份')['每月总人力(USD)'].sum().reset_index()
    avg_burn = actual_burn['每月总人力(USD)'].mean()

    fig_e = go.Figure()
    fig_e.add_trace(go.Scatter(
        x=[actual_burn['月份'].min(), actual_burn['月份'].max() if df_summary.empty else df_summary['月份'].max()],
        y=[avg_burn, avg_burn],
        name=f'历史月均基准 (${avg_burn/1000:.1f}k)',
        line=dict(color='rgba(150,150,150,0.55)', width=2, dash='dot'),
        hoverinfo='skip'
    ))
    fig_e.add_trace(go.Scatter(
        x=actual_burn['月份'],
        y=actual_burn['每月总人力(USD)'],
        name='实际消耗 (Actual)',
        line=dict(color=TECH_COLORS['primary'], width=3, shape='spline'),
        marker=dict(size=10, symbol='circle', color='white', line=dict(width=2, color=TECH_COLORS['primary'])),
        mode='lines+markers+text',
        text=actual_burn['每月总人力(USD)'].apply(lambda x: f"${x/1000:.1f}k"),
        textposition='top center',
        textfont=dict(size=12, color=TECH_COLORS['text_main'])
    ))

    if not df_summary.empty and all(c in df_summary.columns for c in ['固定薪酬', '奖金', '即时激励']):
        last_actual = actual_burn.tail(1)
        df_summary = df_summary.copy()
        df_summary['总成本'] = df_summary['固定薪酬'] + df_summary['奖金'] + df_summary['即时激励']
        forecast_burn = df_summary[['月份', '总成本']].rename(columns={'总成本': '每月总人力(USD)'})
        line_forecast = pd.concat([last_actual, forecast_burn]).sort_values('月份')
        fig_e.add_trace(go.Scatter(
            x=line_forecast['月份'],
            y=line_forecast['每月总人力(USD)'],
            name='未来预测 (Forecast)',
            line=dict(color=TECH_COLORS['primary'], width=3, dash='dash', shape='spline'),
            marker=dict(size=10, symbol='circle', color='white', line=dict(width=2, color=TECH_COLORS['primary'])),
            mode='lines+markers+text',
            text=line_forecast['每月总人力(USD)'].apply(lambda x: f"${x/1000:.1f}k"),
            textposition='top center',
            textfont=dict(size=12, color=TECH_COLORS['text_main']),
            showlegend=True
        ))

    update_fig_layout(fig_e, "", is_cost_chart=True)
    fig_e.update_layout(
        yaxis_title="人力成本消耗 (USD)",
        xaxis=dict(tickformat="%Y-%m", dtick="M1", tickangle=-45)
    )
    st.plotly_chart(fig_e, width="stretch")

with st.expander("明细台账（按需展开）", expanded=False):
    st.dataframe(
        df_filtered.style.format(
            subset=['每月总人力(USD)', '固定薪资(USD)', '奖金/激励(USD)', '即时激励(USD)'],
            formatter="{:,.0f}"
        ),
        use_container_width=True
    )
    csv = df_filtered.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="导出过滤后的数据为 CSV",
        data=csv,
        file_name=f"A65_HR_Data_{datetime.now().strftime('%Y%m%d')}.csv",
        mime='text/csv',
    )
