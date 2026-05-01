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
    initial_sidebar_state="expanded"
)

# --- 互联网大厂高级配色方案 (精准参考参考图配色) ---
TECH_COLORS = {
    'primary': '#5B78D1',    # 参考图“测试组”蓝
    'secondary': '#92CC76',  # 参考图“对照组”绿
    'warning': '#FFC552',    # 参考图“整体”黄
    'danger': '#E86452',     # 柔和红 (保留用于首图变动)
    'neutral': '#5B6370',    # 中性灰
    'background': '#F7F9FC', # 极浅蓝灰背景
    'card_bg': '#FFFFFF',    
    'text_main': '#1F2329',  
    'text_sub': '#646A73'    
}

# --- 样式定制 (平衡专业版：兼顾大字号与布局协调) ---
st.markdown(f"""
    <style>
    /* 全局字体回归平衡：基础 18px */
    html, body, [data-testid="stAppViewContainer"] {{
        font-size: 18px !important;
        background-color: {TECH_COLORS['background']};
        color: {TECH_COLORS['text_main']};
    }}
    
    /* 标题样式：加粗黑色，进一步放大 */
    h1 {{ 
        font-size: 2.2rem !important; 
        font-weight: 800 !important; 
        color: #000000 !important;
        margin-bottom: 0.8rem !important; 
        letter-spacing: -1px;
    }}
    h2 {{ 
        font-size: 1.6rem !important; /* 约 29px */
        font-weight: 900 !important; 
        color: #000000 !important;
        margin-top: 2.5rem !important;
        margin-bottom: 1.2rem !important;
    }}
    h3 {{ 
        font-size: 1.4rem !important; /* 约 25px */
        font-weight: 900 !important; 
        color: #000000 !important; 
        margin-top: 1.5rem !important;
    }}
    
    /* 指标卡回归协调 */
    div[data-testid="stMetric"] {{
        background-color: {TECH_COLORS['card_bg']} !important;
        padding: 20px 25px !important;
        border-radius: 12px !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03) !important;
        border: 1px solid #EFF1F4 !important;
        transition: transform 0.3s ease;
    }}
    div[data-testid="stMetricLabel"] > div {{
        font-size: 1.5rem !important; /* 约 27px */
        font-weight: 900 !important;
        color: #000000 !important; /* 极致加粗黑色 */
    }}
    div[data-testid="stMetricValue"] > div {{
        font-size: 2.2rem !important; /* 约 40px */
        font-weight: 900 !important;
        color: {TECH_COLORS['primary']} !important;
    }}
    
    /* 侧边栏字号同步 */
    section[data-testid="stSidebar"] {{
        width: 360px !important;
    }}
    .stMultiSelect label, .stSelectbox label {{
        font-size: 1.08rem !important;
        font-weight: 700 !important;
    }}
    
    /* 交互组件字号 */
    .stMarkdown p {{
        font-size: 1.08rem !important;
    }}
    
    /* 增强 Plotly 交互提示框 */
    .hoverlayer .hovertext {{
        font-size: 1.0rem !important;
    }}
    
    /* 看数指南 - 恢复浅色底框，便于区分 */
    .stAlert {{
        border: none !important; 
        outline: none !important;
        background-color: #F0F5FF !important; /* 极淡蓝色底框 */
        box-shadow: none !important;
        padding: 0.8rem 1.2rem !important;
        margin-bottom: 1.2rem !important;
        border-radius: 8px !important;
    }}
    .stAlert p {{
        font-size: 1.0rem !important; 
        line-height: 1.5 !important;
        font-weight: 600 !important;
        color: {TECH_COLORS['primary']} !important; 
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

# --- 侧边栏过滤器 (重构：年、季、月多选联动) ---
st.sidebar.header("🔍 数据筛选中心")

# 1. 时间筛选维度
st.sidebar.subheader("📅 时间维度")

# 年份筛选 (多选)
years = sorted(df_detail['年份'].unique().tolist())
sel_years = st.sidebar.multiselect("选择年份", options=years, default=years)

# 基于年份过滤可选季度
if sel_years:
    df_temp_y = df_detail[df_detail['年份'].isin(sel_years)]
    available_quarters = sorted(df_temp_y['季度'].unique().tolist())
else:
    available_quarters = sorted(df_detail['季度'].unique().tolist())
sel_quarters = st.sidebar.multiselect("选择季度", options=[f"Q{q}" for q in available_quarters], default=[f"Q{q}" for q in available_quarters])

# 基于年、季过滤可选月份
mask_m = pd.Series([True] * len(df_detail))
if sel_years:
    mask_m &= (df_detail['年份'].isin(sel_years))
if sel_quarters:
    q_nums = [int(q[1]) for q in sel_quarters]
    mask_m &= (df_detail['季度'].isin(q_nums))

available_months = sorted(df_detail[mask_m]['月'].unique().tolist())
sel_months = st.sidebar.multiselect("选择月份", options=available_months, default=available_months)

# 2. 职能筛选
st.sidebar.subheader("👥 职能分类")
all_categories = sorted(df_detail['职能分类'].unique().tolist())
selected_categories = st.sidebar.multiselect("职能分类", options=all_categories, default=all_categories)

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
st.title("🚀 A65 项目人力成本经营分析看板")
st.markdown(f"**数据状态**：已根据筛选条件加载 {len(df_filtered)} 条记录")

# 顶部指标卡 (增加对比逻辑)
if not df_filtered.empty:
    latest_date = df_filtered['月份'].max()
    curr_data = df_filtered[df_filtered['月份'] == latest_date]
    
    total_hc = len(curr_data)
    total_last_month_cost = curr_data['每月总人力(USD)'].sum()
    total_period_cost = df_filtered['每月总人力(USD)'].sum()
    avg_cost = (total_last_month_cost / total_hc) if total_hc > 0 else 0
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("分析周期总成本", f"${total_period_cost/1000:.1f}k", 
                  help="计算逻辑：当前所选时间范围内所有月份的“每月总人力(USD)”累计总和。")
    with col2:
        st.metric("分析周期末总人数", f"{total_hc} 人", 
                  help="计算逻辑：当前所选时间范围中最后一个月的总人数（Headcount）。")
    with col3:
        st.metric("分析周期末总成本", f"${total_last_month_cost/1000:.1f}k", 
                  help="计算逻辑：对应数据源“明细”表中最后一个月的“每月总人力(USD)”列之和。所有数值以 USD 结算。")
    with col4:
        st.metric("周期末人均单价", f"${avg_cost/1000:.1f}k", 
                  help="计算逻辑：分析周期末总成本 ÷ 分析周期末总人数。")
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
            font=dict(size=25, color=TECH_COLORS['text_main']), 
            x=0, y=0.98
        ),
        template="plotly_white",
        hovermode="closest", 
        clickmode="event+select",
        font=dict(size=17, color=TECH_COLORS['text_sub']), 
        hoverlabel=dict(
            bgcolor="white",
            font_size=18, 
            font_family="PingFang SC, Microsoft YaHei",
            bordercolor=TECH_COLORS['primary'],
            namelength=-1
        ),
        legend=dict(
            orientation="h", 
            yanchor="bottom", y=1.02, 
            xanchor="right", x=1,
            font=dict(size=15) 
        ),
        margin=dict(l=60, r=60, t=100, b=60),
        plot_bgcolor='rgba(0,0,0,0)',
        paper_bgcolor='rgba(0,0,0,0)',
        xaxis=dict(
            showgrid=False, 
            tickfont=dict(size=15),
            title_font=dict(size=17),
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
            tickfont=dict(size=15),
            title_font=dict(size=17),
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

# 1. 月度成本变动监控 (绝对值 + 环比率)
st.subheader("💰 月度成本变动监控 (USD)")
st.markdown(f"""
    <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;'>
        <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
            💡 <b>看数指南</b>：下方柱状图展示<b>本月比上月多花（或少花）了多少钱</b>。红色代表支出增加，绿色代表支出减少。标签显示了变动金额及环比增长率。
        </p>
    </div>
""", unsafe_allow_html=True)

# 计算绝对值差异 (Diff) 和 环比率 (Pct)
grouped_cost = df_detail.groupby('月份').agg({'每月总人力(USD)': 'sum'})
mom_abs_df = grouped_cost.diff().reset_index().rename(columns={'每月总人力(USD)': '成本变动额'})
mom_pct_df = grouped_cost.pct_change().reset_index().rename(columns={'每月总人力(USD)': '环比率'})

# 合并数据
mom_combined = pd.merge(mom_abs_df, mom_pct_df, on='月份').dropna()

# 仅显示过滤范围内的变动
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
    # 清新红绿配色
    marker_color=[TECH_COLORS['danger'] if val >= 0 else TECH_COLORS['secondary'] for val in mom_display['成本变动额']],
    # 强制在每个柱子上显示完整数值(k)和浮动比
    text=mom_display.apply(lambda x: f"<b>{'+' if x['成本变动额'] >= 0 else '-'}${abs(x['成本变动额']/1000):.1f}k</b><br>({x['环比率']:.1%})", axis=1),
    textposition='outside', # 统一放在柱外，避免重叠
    textfont=dict(size=15, color=TECH_COLORS['text_main']),
    cliponaxis=False,
    hovertemplate="月份: %{x|%Y-%m}<br>变动金额: %{y:,.2f} USD<br><extra></extra>"
))

update_fig_layout(fig_mom, "", is_cost_chart=True)
fig_mom.update_layout(
    yaxis=dict(dtick=20000), # 保持 20k 间隔
    xaxis=dict(tickformat="%Y-%m"),
    showlegend=False,
    margin=dict(t=50)
)
st.plotly_chart(fig_mom, use_container_width=True)

# 1.1 变动原因拆解 (变动额的组成结构 + 人数变动)
st.subheader("🔍 月度成本变动归因拆解 (USD)")
st.markdown(f"""
    <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;'>
        <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
            💡 <b>看数指南</b>：此图将每月的总变动额拆解为<b>固定薪资、奖金、即时激励</b>。同时，<b>蓝色折线</b>展示了人数的环比增减。帮助判断：成本上涨是因为调薪，还是因为招人。
        </p>
    </div>
""", unsafe_allow_html=True)

# 计算各分项的月度变动
category_diff = df_detail.groupby('月份').agg({
    '固定薪资(USD)': 'sum',
    '奖金/激励(USD)': 'sum',
    '即时激励(USD)': 'sum',
    '姓名': 'count'
}).diff().reset_index().rename(columns={'姓名': '人数变动'}).dropna()

# 仅显示过滤范围
if not df_filtered.empty:
    diff_display = category_diff[(category_diff['月份'] >= f_min_date) & (category_diff['月份'] <= f_max_date)]
else:
    diff_display = category_diff

# 使用双轴图表
fig_diff_structure = make_subplots(specs=[[{"secondary_y": True}]])

# 添加成本分项变动 (柱状图 - 使用清新配色 + 统一 k 格式)
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

# 添加人数变动 (平滑曲线 + 空心结点样式)
fig_diff_structure.add_trace(
    go.Scatter(
        x=diff_display['月份'], 
        y=diff_display['人数变动'], 
        name='人数变动 (人)', 
        line=dict(color=TECH_COLORS['primary'], width=4, shape='spline'), # 统一为维度D颜色
        marker=dict(size=17, symbol='circle', color='white', line=dict(width=3, color=TECH_COLORS['primary'])), 
        mode='lines+markers+text',
        text=diff_display['人数变动'].apply(lambda x: f"+{int(x)}人" if x > 0 else f"{int(x)}人"),
        textposition='top center',
        textfont=dict(size=15, color=TECH_COLORS['primary'], family="Arial Black") 
    ), 
    secondary_y=True
)

update_fig_layout(fig_diff_structure, "", is_cost_chart=True)
fig_diff_structure.update_layout(
    barmode='relative',
    yaxis_title="成本变动金额 (USD)",
    yaxis2_title="人数变动 (人)",
    xaxis=dict(tickformat="%Y-%m")
)
st.plotly_chart(fig_diff_structure, use_container_width=True)

# 2. 维度 A：人力规模与薪酬结构趋势 (实际 + 预测)
st.subheader("📊 维度 A：人力规模与薪酬结构趋势")
st.markdown(f"""
    <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;'>
        <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
            💡 <b>看数指南</b>：堆积柱状图展示了<b>总成本构成</b>；<b>蓝色折线</b>展示了<b>总人数</b>。虚线/浅色部分代表 2026 年 4 月后的<b>预测数据</b>。
        </p>
    </div>
""", unsafe_allow_html=True)

# --- 数据合并逻辑 ---
# 1. 实际数据汇总
actual_trend = df_filtered.groupby('月份').agg({
    '固定薪资(USD)': 'sum',
    '奖金/激励(USD)': 'sum',
    '即时激励(USD)': 'sum',
    '姓名': 'count'
}).reset_index().rename(columns={
    '固定薪资(USD)': '固定', '奖金/激励(USD)': '奖金', '即时激励(USD)': '即时激励', '姓名': '人数'
})
actual_trend['类型'] = '实际'
actual_trend['总成本'] = actual_trend['固定'] + actual_trend['奖金'] + actual_trend['即时激励']

# 2. 预测数据处理 (从 df_summary 提取)
forecast_trend = pd.DataFrame()
if not df_summary.empty and '月份' in df_summary.columns:
    # 筛选出 2026-04 以后的预测数据
    f_mask = (df_summary['月份'] >= '2026-04-01')
    
    # 检查是否有列
    required_f_cols = ['固定薪酬', '奖金', '即时激励', '人数']
    existing_f_cols = [c for c in required_f_cols if c in df_summary.columns]
    
    if existing_f_cols:
        forecast_trend = df_summary[f_mask].copy()
        if not forecast_trend.empty:
            forecast_trend = forecast_trend.rename(columns={
                '固定薪酬': '固定', '奖金': '奖金', '即时激励': '即时激励'
            })
            forecast_trend['类型'] = '预测'
            # 确保人数和成本列是数值
            for c in ['人数', '固定', '奖金', '即时激励']:
                if c in forecast_trend.columns:
                    forecast_trend[c] = pd.to_numeric(forecast_trend[c], errors='coerce').fillna(0)
            forecast_trend['总成本'] = forecast_trend['固定'] + forecast_trend['奖金'] + forecast_trend['即时激励']

fig_combined = make_subplots(specs=[[{"secondary_y": True}]])

# --- 绘制实际数据 (实色 - 使用新配色 + 统一 k 格式) ---
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

# --- 绘制预测数据 (浅色/条纹 - 隐藏图例 + 统一 k 格式) ---
if not forecast_trend.empty:
    fig_combined.add_trace(go.Bar(
        x=forecast_trend['月份'], y=forecast_trend['固定'], name='固定薪资 (预测)', marker_color=TECH_COLORS['primary'], 
        opacity=0.4, marker_pattern_shape="/", showlegend=False,
        text=forecast_trend['固定'].apply(lambda x: f"${x/1000:.1f}k"), textposition='inside'
    ), secondary_y=False)
    fig_combined.add_trace(go.Bar(
        x=forecast_trend['月份'], y=forecast_trend['奖金'], name='奖金/激励 (预测)', marker_color=TECH_COLORS['secondary'], 
        opacity=0.4, marker_pattern_shape="/", showlegend=False,
        text=forecast_trend['奖金'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
    ), secondary_y=False)
    fig_combined.add_trace(go.Bar(
        x=forecast_trend['月份'], y=forecast_trend['即时激励'], name='即时激励 (预测)', marker_color=TECH_COLORS['warning'], 
        opacity=0.4, marker_pattern_shape="/", showlegend=False,
        text=forecast_trend['即时激励'].apply(lambda x: f"${x/1000:.1f}k" if x > 0 else ""), textposition='inside'
    ), secondary_y=False)

# 添加总成本标签 (维度A)
total_trend = pd.concat([actual_trend, forecast_trend]).sort_values('月份')
fig_combined.add_trace(go.Scatter(
    x=total_trend['月份'], y=total_trend['总成本'],
    mode='text',
    text=total_trend['总成本'].apply(lambda x: f"<b>${x/1000:.1f}k</b>"),
    textposition='top center',
    textfont=dict(size=14, color=TECH_COLORS['text_main']),
    showlegend=False,
    hoverinfo='skip'
), secondary_y=False)

# --- 绘制人数折线 (实际为实线，预测为虚线) ---
# 合并实际和预测的人数数据用于画线
full_hc = pd.concat([actual_trend[['月份', '人数', '类型']], forecast_trend[['月份', '人数', '类型']]]).sort_values('月份')

# 实际人数线 (平滑 + 空心结点)
fig_combined.add_trace(
    go.Scatter(
        x=actual_trend['月份'], y=actual_trend['人数'], 
        name="总人数 (实际)", 
        line=dict(color=TECH_COLORS['primary'], width=4, shape='spline'), # 统一为维度D颜色
        marker=dict(size=17, symbol='circle', color='white', line=dict(width=3, color=TECH_COLORS['primary'])),
        mode='lines+markers+text',
        text=actual_trend['人数'].apply(lambda x: f"{int(x)}人"),
        textposition='top center',
        textfont=dict(size=15, color=TECH_COLORS['primary'], family="Arial Black") 
    ),
    secondary_y=True
)

# 预测人数线 (虚线 + 平滑 + 空心结点 - 隐藏图例)
if not forecast_trend.empty:
    last_actual = actual_trend.tail(1)
    line_forecast = pd.concat([last_actual, forecast_trend]).sort_values('月份')
    
    fig_combined.add_trace(
        go.Scatter(
            x=line_forecast['月份'], y=line_forecast['人数'], 
            name="总人数 (预测)", 
            line=dict(color=TECH_COLORS['primary'], width=4, dash='dash', shape='spline'), # 统一为维度D颜色
            marker=dict(size=17, symbol='circle', color='white', line=dict(width=3, color=TECH_COLORS['primary'])),
            mode='lines+markers+text', # 恢复 +text
            text=line_forecast['人数'].apply(lambda x: f"{int(x)}人"),
            textposition='top center',
            textfont=dict(size=15, color=TECH_COLORS['primary'], family="Arial Black"),
            showlegend=False # 隐藏图例
        ),
        secondary_y=True
    )

# --- 绘制总成本信息 (仅用于 Hover 提示，不显示标签) ---
fig_combined.add_trace(
    go.Scatter(
        x=actual_trend['月份'], 
        y=actual_trend['总成本'], 
        name='<b>总成本 (实际)</b>',
        mode='markers',
        marker=dict(opacity=0), # 隐藏结点
        hovertemplate="总成本: $%{y:,.2f}<extra></extra>",
        showlegend=False
    ),
    secondary_y=False
)

if not forecast_trend.empty:
    fig_combined.add_trace(
        go.Scatter(
            x=forecast_trend['月份'], 
            y=forecast_trend['总成本'], 
            name='<b>总成本 (预测)</b>',
            mode='markers',
            marker=dict(opacity=0), # 隐藏结点
            hovertemplate="总成本: $%{y:,.2f}<extra></extra>",
            showlegend=False
        ),
        secondary_y=False
    )

update_fig_layout(fig_combined, "", is_cost_chart=True)
fig_combined.update_layout(
    barmode='stack', 
    yaxis_title="人力成本 (USD)",
    yaxis2_title="人数 (人)",
    xaxis=dict(tickformat="%Y-%m")
)
st.plotly_chart(fig_combined, use_container_width=True)

# 3. 维度 B & D
c1, c2 = st.columns(2)
with c1:
    # 维度 B：职能与岗位结构深度穿透 (分析周期末快照)
    st.subheader("维度 B：职能与岗位结构深度穿透")
    
    # 逻辑修正：仅取筛选范围内的【最后一个月】的数据，展示实时结构，而非累计人次
    if not df_filtered.empty:
        latest_month_in_filter = df_filtered['月份'].max()
        sun_data_latest = df_filtered[df_filtered['月份'] == latest_month_in_filter]
        
        st.markdown(f"""
            <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;'>
                <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
                    💡 <b>看数指南</b>：此图展示 <b>{latest_month_in_filter.strftime('%Y-%m')}</b> 的人员结构快照。内环为职能，外环为岗位性质。点击内环可钻取。
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        # 准备旭日图数据
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
            font=dict(size=15),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        st.plotly_chart(fig_b, use_container_width=True)
    else:
        st.warning("无数据可展示")

with c2:
    st.subheader("维度 C：职能人均成本构成")
    
    # 方案 C：分项人均逻辑
    # 1. 按职能汇总各项总额和总人次
    unit_breakdown = df_filtered.groupby('职能分类').agg({
        '固定薪资(USD)': 'sum',
        '奖金/激励(USD)': 'sum',
        '即时激励(USD)': 'sum',
        '姓名': 'count'
    }).reset_index()

    # 2. 计算各项的人均值
    unit_breakdown['人均固定'] = unit_breakdown['固定薪资(USD)'] / unit_breakdown['姓名']
    unit_breakdown['人均奖金'] = unit_breakdown['奖金/激励(USD)'] / unit_breakdown['姓名']
    unit_breakdown['人均即时'] = unit_breakdown['即时激励(USD)'] / unit_breakdown['姓名']
    unit_breakdown['总人均'] = unit_breakdown['人均固定'] + unit_breakdown['人均奖金'] + unit_breakdown['人均即时']
    
    # 排序
    unit_breakdown = unit_breakdown.sort_values('总人均', ascending=False)

    with st.expander("📄 查看维度 C 计算规则", expanded=False):
        st.markdown(f"""
            <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 15px;'>
                <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
                    💡 <b>看数指南</b>：此图展示各职能平均每人每月成本构成。可以直观看到：人均成本高是因为<b>底薪高</b>还是<b>奖金多</b>。
                </p>
            </div>
        """, unsafe_allow_html=True)
        st.write("""
        **计算逻辑：方案 C (职能维度汇总)**
        1. **分项汇总**：将所选时间段内，各职能的“固定薪资”、“奖金”、“即时激励”分别求和。
        2. **人数汇总**：统计该职能下的“总人次”（即每个月的人数加总）。
        3. **人均计算**：分项总额 / 总人次。
        *注：此逻辑反映了该职能在该期间内的平均投入水平，不受月份长短波动影响。*
        
        ---
        ### **📊 原始数据核对清单**
        """)
        check_df = unit_breakdown[['职能分类', '姓名', '固定薪资(USD)', '奖金/激励(USD)', '总人均']].copy()
        check_df.columns = ['职能分类', '累计总人次(月/人)', '固定薪资累计', '奖金/激励累计', '计算出的人均月成本']
        st.dataframe(check_df.style.format({
            '固定薪资累计': '${:,.2f}',
            '奖金/激励累计': '${:,.2f}',
            '计算出的人均月成本': '${:,.2f}'
        }), use_container_width=True)
        
        st.caption("注：如果某员工月中入职导致当月工资未满额，仍会被计为1人次，这会拉低该岗位的平均值。")

    fig_c_stack = go.Figure()
    # 添加堆叠分项 (使用新配色)
    fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均固定'], name='人均固定薪资', marker_color=TECH_COLORS['primary']))
    fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均奖金'], name='人均奖金/激励', marker_color=TECH_COLORS['secondary']))
    fig_c_stack.add_trace(go.Bar(x=unit_breakdown['职能分类'], y=unit_breakdown['人均即时'], name='人均即时激励', marker_color=TECH_COLORS['warning']))

    # 添加总计标签 (增大字体 + 统一 k 格式)
    fig_c_stack.add_trace(go.Scatter(
        x=unit_breakdown['职能分类'], 
        y=unit_breakdown['总人均'], 
        mode='text',
        text=unit_breakdown['总人均'].apply(lambda x: f"<b>${x/1000:.1f}k</b>"),
        textposition='top center',
        textfont=dict(size=15, color=TECH_COLORS['text_main']), 
        showlegend=False,
        hoverinfo='skip'
    ))

    update_fig_layout(fig_c_stack, "", is_cost_chart=True)
    fig_c_stack.update_layout(
        barmode='stack',
        yaxis_title="人均成本 (USD/月)",
        xaxis=dict(tickformat=None, dtick=None),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
    )
    st.plotly_chart(fig_c_stack, use_container_width=True)

# 4. 预测对比
st.subheader("维度 D：预测 vs. 实际 (Burn Rate)")
st.markdown(f"""
    <div style='background-color: #F0F5FF; padding: 12px 20px; border-radius: 8px; margin-bottom: 20px;'>
        <p style='font-size: 1.0rem; font-weight: 600; color: {TECH_COLORS['primary']}; margin: 0;'>
            💡 <b>看数指南</b>：此图对比了每月总成本的<b>实际消耗</b>与<b>未来预测</b>。<br>
            1. <b>水平基准线</b>：由历史月均 Burn Rate 自动计算的基准。<br>
            2. <b>蓝色实线</b>：历史实际发生的每月总人力成本。<br>
            3. <b>蓝色虚线</b>：基于“预测明细”计算的未来各月预计成本。
        </p>
    </div>
""", unsafe_allow_html=True)
actual_burn = df_detail.groupby('月份')['每月总人力(USD)'].sum().reset_index()

# 计算历史月均 Burn Rate (基准线)
avg_burn = actual_burn['每月总人力(USD)'].mean()

fig_e = go.Figure()

# 1. 绘制历史月均基准线
fig_e.add_trace(go.Scatter(
    x=[actual_burn['月份'].min(), actual_burn['月份'].max() if df_summary.empty else df_summary['月份'].max()],
    y=[avg_burn, avg_burn],
    name=f'历史月均基准 (${avg_burn:,.2f})',
    line=dict(color='rgba(150, 150, 150, 0.5)', width=2, dash='dot'),
    hoverinfo='skip'
))

# 2. 绘制实际消耗 (实线 - 使用清新蓝 + 统一 k 格式)
fig_e.add_trace(go.Scatter(
    x=actual_burn['月份'], 
    y=actual_burn['每月总人力(USD)'], 
    name='实际消耗 (Actual)', 
    line=dict(color=TECH_COLORS['primary'], width=5, shape='spline'),
    marker=dict(size=17, symbol='circle', color='white', line=dict(width=3, color=TECH_COLORS['primary'])),
    mode='lines+markers+text',
    text=actual_burn['每月总人力(USD)'].apply(lambda x: f"${x/1000:.1f}k"),
    textposition='top center',
    textfont=dict(size=15, color=TECH_COLORS['text_main'], family="Arial Black") 
))

# 3. 绘制预测消耗 (虚线 - 使用清新蓝 + 统一 k 格式)
if not df_summary.empty:
    # 为了线段连续，包含最后一个实际点
    last_actual = actual_burn.tail(1)
    # 计算预测表的总成本 (固定+奖金+即时)
    df_summary['总成本'] = df_summary['固定薪酬'] + df_summary['奖金'] + df_summary['即时激励']
    forecast_burn = df_summary[['月份', '总成本']].rename(columns={'总成本': '每月总人力(USD)'})
    line_forecast = pd.concat([last_actual, forecast_burn]).sort_values('月份')
    
    fig_e.add_trace(go.Scatter(
        x=line_forecast['月份'], 
        y=line_forecast['每月总人力(USD)'], 
        name='未来预测 (Forecast)', 
        line=dict(color=TECH_COLORS['primary'], width=5, dash='dash', shape='spline'),
        marker=dict(size=17, symbol='circle', color='white', line=dict(width=3, color=TECH_COLORS['primary'])),
        mode='lines+markers+text',
        text=line_forecast['每月总人力(USD)'].apply(lambda x: f"${x/1000:.1f}k"),
        textposition='top center',
        textfont=dict(size=15, color=TECH_COLORS['text_main'], family="Arial Black") 
    ))

update_fig_layout(fig_e, "", is_cost_chart=True)
fig_e.update_layout(yaxis_title="人力成本消耗 (USD)")
st.plotly_chart(fig_e, use_container_width=True)

# 数据导出
st.markdown("---")
st.subheader("📂 数据清单与导出")
st.dataframe(df_filtered.style.format(subset=['每月总人力(USD)', '固定薪资(USD)', '奖金/激励(USD)'], formatter="{:.2f}"), use_container_width=True)

csv = df_filtered.to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 导出过滤后的数据为 CSV",
    data=csv,
    file_name=f"A65_HR_Data_{datetime.now().strftime('%Y%m%d')}.csv",
    mime='text/csv',
)

st.caption("Developed by AI Assistant | 互联网大厂风格经营分析系统")
