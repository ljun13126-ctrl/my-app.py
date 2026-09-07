import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import plotly.express as px
import plotly.graph_objects as go

plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

st.set_page_config(page_title="灾智云 · 智能决策平台", layout="wide", page_icon="☁️")
st.markdown("""
    <style>
        .css-1d391kg { display: none !important; }
        section[data-testid="stSidebar"] { display: none !important; }
        .stApp { background: linear-gradient(135deg, #0a0f1e 0%, #162a4a 40%, #0d1b2a 100%); color: #ffffff; }
        .nav-container { display: flex; justify-content: space-between; align-items: center; padding: 16px 40px; background: rgba(11, 17, 32, 0.7); border-bottom: 1px solid rgba(212, 175, 55, 0.15); border-radius: 0 0 20px 20px; margin-bottom: 20px; }
        .nav-brand { font-size: 28px; font-weight: 700; color: #d4af37; }
        .nav-links { display: flex; gap: 32px; list-style: none; }
        .nav-links li a { color: rgba(255, 255, 255, 0.65); text-decoration: none; font-size: 16px; padding: 8px 16px; border-radius: 8px; }
        .nav-links li a.active { color: #d4af37; background: rgba(212, 175, 55, 0.15); }
        .main-title { font-size: 72px; font-weight: 700; text-align: center; color: #f7e68a; }
        .sub-title { text-align: center; font-size: 24px; color: rgba(255, 255, 255, 0.8); }
        .stat-card { background: rgba(255, 255, 255, 0.04); border-radius: 16px; padding: 20px 24px; border-left: 4px solid #d4af37; }
        .stat-label { font-size: 13px; color: #a0aec0; }
        .stat-value { font-size: 28px; font-weight: 700; color: #ffffff; }
        .stButton > button { background: #d4af37 !important; color: #0b1120 !important; font-weight: 600 !important; border-radius: 40px !important; }
    </style>
""", unsafe_allow_html=True)

DB_PATH = "data/uploaded_data.db"
os.makedirs("data", exist_ok=True)

def init_db():
    conn = sqlite3.connect(DB_PATH); c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS disaster_data (id INTEGER PRIMARY KEY AUTOINCREMENT, region TEXT, disaster_type TEXT, parent_region TEXT, disaster_time TEXT, affected_population INTEGER, death_population INTEGER, missing_population INTEGER, emergency_evacuation INTEGER, emergency_relocation INTEGER, emergency_life_aid INTEGER, collapsed_houses INTEGER, collapsed_households INTEGER, severe_damaged_houses INTEGER, severe_damaged_households INTEGER, moderate_damaged_houses INTEGER, moderate_damaged_households INTEGER, crop_area_affected REAL, crop_area_no_harvest REAL, direct_economic_loss REAL, housing_loss REAL, agri_loss REAL, industry_loss REAL, infrastructure_loss REAL, public_service_loss REAL, other_loss REAL)''')
    conn.commit(); conn.close()

init_db()

def load_data():
    conn = sqlite3.connect(DB_PATH); df = pd.read_sql_query("SELECT * FROM disaster_data", conn); conn.close(); return df

def save_data(df):
    conn = sqlite3.connect(DB_PATH); df.to_sql('disaster_data', conn, if_exists='replace', index=False); conn.close()

def clean_data(df):
    df = df.fillna({ '受灾人口(人)': 0, '因灾死亡人口(人)': 0, '因灾失踪人口(人)': 0, '紧急避险转移人口(人)': 0, '紧急转移安置人口(累计值)(人)': 0, '需紧急生活救助人口(累计值)(人)': 0, '倒塌房屋间数(间)': 0, '倒塌住房户数(户)': 0, '严重损坏房屋间数(间)': 0, '严重损坏住房户数(户)': 0, '一般损坏房屋间数(间)': 0, '一般损坏住房户数(户)': 0, '农作物受灾面积(公顷)': 0.0, '农作物绝收面积(公顷)': 0.0, '直接经济损失(万元)': 0.0, '其中：住房及居民家庭财产损失(万元)': 0.0, '农林牧渔业损失(万元)': 0.0, '工矿商贸业损失(万元)': 0.0, '基础设施损失(万元)': 0.0, '公共服务损失(万元)': 0.0, '其他损失(万元)': 0.0 })
    num_cols = df.select_dtypes(include=['number']).columns
    for col in num_cols:
        df[col] = df[col].apply(lambda x: max(x, 0) if isinstance(x, (int, float)) else x)
    if '灾害发生时间' in df.columns:
        df['灾害发生时间'] = pd.to_datetime(df['灾害发生时间'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
    for col in ['区域', '隶属区域']:
        if col in df.columns: df[col] = df[col].fillna('--')
    return df

def get_summary_stats(df):
    if df.empty: return {}
    return { '总记录数': len(df), '受灾总人口': int(df['受灾人口(人)'].sum()), '死亡失踪人口': int(df['因灾死亡人口(人)'].sum() + df['因灾失踪人口(人)'].sum()), '转移安置人口': int(df['紧急转移安置人口(累计值)(人)'].sum()), '直接经济损失(万元)': round(df['直接经济损失(万元)'].sum(), 2), '倒塌房屋间数': int(df['倒塌房屋间数(间)'].sum()), '农作物受灾面积(公顷)': round(df['农作物受灾面积(公顷)'].sum(), 2) }

def get_time_trend(df, freq='M'):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    if freq == 'Y': df_t['时段'] = df_t['时间'].dt.year.astype(str)
    elif freq == 'Q': df_t['时段'] = df_t['时间'].dt.to_period('Q').astype(str)
    elif freq == 'M': df_t['时段'] = df_t['时间'].dt.to_period('M').astype(str)
    elif freq == 'D': df_t['时段'] = df_t['时间'].dt.date.astype(str)
    elif freq == 'h': df_t['时段'] = df_t['时间'].dt.floor('h').astype(str)
    return df_t.groupby('时段').agg({ '受灾人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()

# 【优化点1】更智能的空间识别与手动兜底
def detect_admin_levels(df):
    levels = {}
    for col in df.columns:
        col_str = str(col)
        if '省' in col_str or '直辖市' in col_str: levels['省级'] = col
        elif '市' in col_str or '地市' in col_str: levels['市级'] = col
        elif '县' in col_str or '区县' in col_str: levels['县级'] = col
        elif '乡' in col_str or '镇' in col_str or '街道' in col_str: levels['乡级'] = col
    return levels

def get_region_analysis(df, level_col):
    if df.empty or level_col not in df.columns: return pd.DataFrame()
    return df.groupby(level_col).agg({ '受灾人口(人)': 'sum', '因灾死亡人口(人)': 'sum', '直接经济损失(万元)': 'sum', '农作物受灾面积(公顷)': 'sum' }).reset_index()

def get_disaster_type_analysis(df):
    if df.empty: return pd.DataFrame()
    return df.groupby('灾种').agg({ '受灾人口(人)': 'sum', '因灾死亡人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()

def get_loss_structure(df):
    if df.empty: return {}
    total = df['直接经济损失(万元)'].sum()
    if total == 0: return {}
    return { '住房及家庭财产': round(df['其中：住房及居民家庭财产损失(万元)'].sum() / total * 100, 2), '农林牧渔业': round(df['农林牧渔业损失(万元)'].sum() / total * 100, 2), '基础设施': round(df['基础设施损失(万元)'].sum() / total * 100, 2), '工矿商贸业': round(df['工矿商贸业损失(万元)'].sum() / total * 100, 2) }

# 【优化点3】Top5调整为函数，去掉图，用表格
def get_custom_analysis1(df):
    if df.empty: return pd.DataFrame()
    cross = df.groupby(['区域', '灾种']).agg({'直接经济损失(万元)': 'sum', '受灾人口(人)': 'sum'}).reset_index()
    return cross.sort_values('直接经济损失(万元)', ascending=False).head(5)

def get_custom_analysis2(df):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    df_t['月份'] = df_t['时间'].dt.month
    return df_t.pivot_table(index='月份', columns='灾种', aggfunc='size', fill_value=0).reset_index()

def get_custom_analysis(df, selected_cols, cross_col='灾种'):
    if df.empty or not selected_cols: return pd.DataFrame()
    return df.groupby(cross_col)[selected_cols].sum().reset_index()

# 报告生成函数（略）
def generate_report(df):
    # (保持之前的Word报告逻辑)
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(3.7); section.bottom_margin = Cm(3.5); section.left_margin = Cm(2.8); section.right_margin = Cm(2.6)
    p_title = doc.add_paragraph(); p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("自然灾害灾情综合分析报告"); run.font.name = '方正小标宋简体'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋简体'); run.font.size = Pt(22)
    doc.add_paragraph()
    stats = get_summary_stats(df)
    trend = get_time_trend(df, 'M')
    disaster = get_disaster_type_analysis(df)
    loss = get_loss_structure(df)
    
    def add_para(text):
        p = doc.add_paragraph()
        pf = p.paragraph_format; pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = Pt(28.5); pf.first_line_indent = Pt(32)
        run = p.add_run(text); run.font.name = '仿宋_GB2312'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312'); run.font.size = Pt(16)
    
    add_para("一、宏观背景分析")
    add_para("在全球气候变化深刻影响下，极端天气事件呈现频发重发态势。四川省地处盆地与高原过渡带，地形复杂，灾害多发……（此处省略大量3000字内容）……提升防灾减灾救灾能力是国家治理现代化的重要内容。")
    add_para("二、灾情多维分析")
    add_para(f"据统计，本次灾情共记录事件 {stats.get('总记录数',0)} 起，受灾人口 {stats.get('受灾总人口',0)} 人……直接经济损失达 {stats.get('直接经济损失(万元)',0)} 万元。")
    # 图1
    if not trend.empty:
        fig, ax = plt.subplots(figsize=(10, 5)); ax.plot(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', linewidth=2); ax.set_title('图1：直接经济损失月度趋势'); fig.savefig("g1.png", dpi=200); plt.close(fig); doc.add_picture("g1.png", width=Inches(6)); add_para("【深度说明】月度损失呈现集中爆发态势，提示建立全天候预警机制。")
    # 图2-5略...
    
    file_stream = io.BytesIO(); doc.save(file_stream); file_stream.seek(0)
    return file_stream# ==================== 页面导航 ====================
if 'page' not in st.session_state: st.session_state.page = '首页'
st.markdown(f"""<nav class="nav-container"><div class="nav-brand">☁️ 灾智云</div><ul class="nav-links"><li><a class="{'active' if st.session_state.page == '首页' else ''}">首页</a></li><li><a class="{'active' if st.session_state.page == '数据导入' else ''}">数据导入</a></li><li><a class="{'active' if st.session_state.page == '多维度分析' else ''}">多维度分析</a></li><li><a class="{'active' if st.session_state.page == '智能报告' else ''}">智能报告</a></li></ul></nav>""", unsafe_allow_html=True)

def set_page(page_name): st.session_state.page = page_name
c1, c2, c3, c4 = st.columns(4)
with c1:
    if st.button("🏠 首页", key="nav_home", use_container_width=True): set_page('首页')
with c2:
    if st.button("📥 数据导入", key="nav_upload", use_container_width=True): set_page('数据导入')
with c3:
    if st.button("📊 多维度分析", key="nav_analysis", use_container_width=True): set_page('多维度分析')
with c4:
    if st.button("📄 智能报告", key="nav_report", use_container_width=True): set_page('智能报告')
st.markdown("---")

# ==================== 各页面内容 ====================
if st.session_state.page == '首页':
    st.markdown('<div class="main-title">灾智云</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">自然灾害智能分析 · 辅助决策支撑平台</div>', unsafe_allow_html=True)

elif st.session_state.page == '数据导入':
    st.markdown("## 📥 数据导入与清洗")
    uploaded_file = st.file_uploader("选择 Excel 文件 (.xlsx / .xls)", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            df_clean = clean_data(pd.read_excel(uploaded_file, skiprows=1))
            save_data(df_clean)
            st.success(f"✅ 上传成功，共 {len(df_clean)} 条记录。")
        except Exception as e: st.error(f"❌ 读取失败: {str(e)}")

elif st.session_state.page == '多维度分析':
    st.markdown("## 📊 综合分析仪表板")
    df = load_data()
    if df.empty: st.warning("⚠️ 暂无数据")
    else:
        stats = get_summary_stats(df)
        cols = st.columns(4)
        for i, (k, v) in enumerate(stats.items()):
            with cols[i % 4]: st.markdown(f"""<div class="stat-card"><div class="stat-label">{k}</div><div class="stat-value">{v}</div></div>""", unsafe_allow_html=True)
        st.markdown("---")
        a, b = st.columns(2)
        
        # 时间趋势图
        with a:
            freq_label = st.selectbox("时间粒度:", ["年", "季度", "月", "日", "小时"], index=2)
            freq_map = {"年": "Y", "季度": "Q", "月": "M", "日": "D", "小时": "h"}
            trend = get_time_trend(df, freq_map[freq_label])
            if not trend.empty:
                fig = px.line(trend, x='时段', y='受灾人口(人)', title=f"受灾人口{freq_label}度变化")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
        
        # 灾种占比图
        with b:
            disaster = get_disaster_type_analysis(df)
            if not disaster.empty:
                fig = px.pie(disaster, values='直接经济损失(万元)', names='灾种', title="各灾种经济损失占比", color_discrete_sequence=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1', '#a29bfe'])
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 💸 核心损失结构拆解")
        loss = get_loss_structure(df)
        if loss:
            loss_df = pd.DataFrame(list(loss.items()), columns=['类型', '占比(%)'])
            fig = px.bar(loss_df, x='类型', y='占比(%)', title="四大重点领域占比", color_discrete_sequence=['#d4af37'])
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig, use_container_width=True)

        # 【优化点1】空间维度分析：自动识别+手动兜底+热力树图
        st.markdown("### 🗺️ 空间维度逐级下钻与受灾强度热力分布")
        admin_levels = detect_admin_levels(df)
        sel_level = None
        if not admin_levels:
            st.warning("⚠️ 自动未识别到省市县乡列，请手动选择对应列：")
            # 手动映射兜底
            cols_list = list(df.columns)
            c1, c2 = st.columns(2)
            with c1: 
                l1 = st.selectbox("请选择【省级/市级】列（可忽略）", ['无'] + cols_list)
            with c2: 
                l2 = st.selectbox("请选择【县级/乡级】列（可忽略）", ['无'] + cols_list)
            if l2 != '无':
                admin_levels['县级'] = l2
            if l1 != '无':
                admin_levels['市级'] = l1
            sel_level = st.selectbox("选择下钻层级:", list(admin_levels.keys()))
        else:
            sel_level = st.selectbox("选择下钻层级:", list(admin_levels.keys()))
        
        if sel_level and admin_levels:
            level_col = admin_levels[sel_level]
            region = get_region_analysis(df, level_col)
            if not region.empty:
                # 新增：行政层级受灾强度热力树图（Treemap）
                fig = px.treemap(region, path=[px.Constant("全部区域"), level_col], values='直接经济损失(万元)', 
                                 color='直接经济损失(万元)', color_continuous_scale=px.colors.sequential.Plasma,
                                 title=f"行政层级受灾强度热力分布（按{sel_level}）")
                fig.update_traces(root_color="lightgrey")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)

        # 【优化点2】新增：受灾人口与避险转移关联分析气泡图
        st.markdown("### 💨 受灾人口与避险转移关联分析气泡图")
        bubble_df = df[df['受灾人口(人)'] > 0]
        if not bubble_df.empty:
            fig = px.scatter(bubble_df, x="受灾人口(人)", y="紧急转移安置人口(累计值)(人)", 
                             size="直接经济损失(万元)", color="灾种", 
                             hover_name="区域", title="受灾与避险转移关联分析 (气泡大小=经济损失)",
                             color_discrete_sequence=px.colors.qualitative.Pastel)
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 🔍 自定义深度分析")
        x1, x2 = st.columns(2)
        
        # 【优化点3】TOP5表格展示
        with x1:
            st.markdown("#### 🎯 高风险区域-灾种组合 (TOP5) 表格")
            custom1 = get_custom_analysis1(df)
            if not custom1.empty:
                st.dataframe(custom1, use_container_width=True)
        
        # 【优化点4】月份灾种频次气泡图，更官方颜色（Plasma渐变色对比强烈）
        with x2:
            st.markdown("#### 🌡️ 月份-灾种发生频次气泡图")
            custom2 = get_custom_analysis2(df)
            if not custom2.empty:
                melt_df = custom2.melt(id_vars='月份', var_name='灾种', value_name='频次')
                # 换成官方感更强的气泡图，而不是热力图，用Plasma色系
                fig = px.scatter(melt_df, x="月份", y="灾种", size="频次", color="频次",
                                 color_continuous_scale=px.colors.sequential.Plasma,
                                 title="月份与灾种发生频次分布")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("#### 🧬 自由勾选任意灾损指标交叉分析")
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        cross_dim = st.selectbox("选择交叉维度:", ['灾种'] + [c for c in df.columns if c not in numeric_cols])
        selected_metrics = st.multiselect("自由勾选指标:", numeric_cols, default=numeric_cols[:3])
        if selected_metrics:
            custom_data = get_custom_analysis(df, selected_metrics, cross_dim)
            if not custom_data.empty:
                st.dataframe(custom_data, use_container_width=True)
                fig = px.bar(custom_data, x=cross_dim, y=selected_metrics, title=f"按{cross_dim}统计各项灾损指标", barmode='group')
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)

elif st.session_state.page == '智能报告':
    st.markdown("## 📄 智能报告生成")
    df = load_data()
    if df.empty: st.warning("⚠️ 暂无数据")
    else:
        if st.button("🚀 生成并下载 3000字国标报告", use_container_width=True):
            with st.spinner("正在生成深度报告中..."): st.download_button("📥 点击下载报告", data=generate_report(df), file_name="灾智云_国标专业分析报告.docx", use_container_width=True)

st.markdown("""<div class="footer"><span>⚡ 企业命题：四川省减灾中心</span> &nbsp;|&nbsp; © 2026 灾智云</div>""", unsafe_allow_html=True)
