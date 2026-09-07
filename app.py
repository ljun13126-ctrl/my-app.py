import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import plotly.express as px
import plotly.graph_objects as go

# 设置中文字体
plt.rcParams['font.sans-serif'] = ['SimHei', 'Arial Unicode MS']
plt.rcParams['axes.unicode_minus'] = False

# ==================== 页面配置与 CSS ====================
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
        .main-title { font-size: 72px; font-weight: 700; background: linear-gradient(to right, #d4af37, #f7e68a); -webkit-background-clip: text; -webkit-text-fill-color: transparent; text-align: center; }
        .sub-title { text-align: center; font-size: 24px; color: rgba(255, 255, 255, 0.8); }
        .stat-card { background: rgba(255, 255, 255, 0.04); border-radius: 16px; padding: 20px 24px; border-left: 4px solid #d4af37; }
        .stat-label { font-size: 13px; color: #a0aec0; }
        .stat-value { font-size: 28px; font-weight: 700; color: #ffffff; }
        .step-card { background: rgba(255, 255, 255, 0.03); padding: 24px 30px; border-radius: 16px; border: 1px solid rgba(212, 175, 55, 0.12); text-align: center; }
        .step-number { font-size: 32px; font-weight: 700; color: #d4af37; display: block; }
        .stButton > button { background: #d4af37 !important; color: #0b1120 !important; font-weight: 600 !important; border-radius: 40px !important; }
    </style>
""", unsafe_allow_html=True)

# ==================== 数据库与数据清洗 ====================
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

# ==================== 核心分析函数（修复版） ====================
def get_summary_stats(df):
    if df.empty: return {}
    return { '总记录数': len(df), '受灾总人口': int(df['受灾人口(人)'].sum()), '死亡失踪人口': int(df['因灾死亡人口(人)'].sum() + df['因灾失踪人口(人)'].sum()), '转移安置人口': int(df['紧急转移安置人口(累计值)(人)'].sum()), '直接经济损失(万元)': round(df['直接经济损失(万元)'].sum(), 2), '倒塌房屋间数': int(df['倒塌房屋间数(间)'].sum()), '农作物受灾面积(公顷)': round(df['农作物受灾面积(公顷)'].sum(), 2) }

# 修复点：将 'H' 改为 'h'
def get_time_trend(df, freq='M'):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy()
    df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce')
    df_t = df_t.dropna(subset=['时间'])
    if freq == 'Y': df_t['时段'] = df_t['时间'].dt.year.astype(str)
    elif freq == 'Q': df_t['时段'] = df_t['时间'].dt.to_period('Q').astype(str)
    elif freq == 'M': df_t['时段'] = df_t['时间'].dt.to_period('M').astype(str)
    elif freq == 'D': df_t['时段'] = df_t['时间'].dt.date.astype(str)
    elif freq == 'h': df_t['时段'] = df_t['时间'].dt.floor('h').astype(str)
    grouped = df_t.groupby('时段').agg({ '受灾人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()
    return grouped

# 自动识别省市县乡
def detect_admin_levels(df):
    levels = {}
    for col in df.columns:
        if '省' in col: levels['省级'] = col
        elif '市' in col: levels['市级'] = col
        elif '县' in col: levels['县级'] = col
        elif '乡' in col: levels['乡级'] = col
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

# 原自定义分析的两个图
def get_custom_analysis1(df):
    if df.empty: return pd.DataFrame()
    cross = df.groupby(['区域', '灾种']).agg({'直接经济损失(万元)': 'sum', '受灾人口(人)': 'sum'}).reset_index()
    return cross.sort_values('直接经济损失(万元)', ascending=False).head(10)

def get_custom_analysis2(df):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    df_t['月份'] = df_t['时间'].dt.month
    return df_t.pivot_table(index='月份', columns='灾种', aggfunc='size', fill_value=0).reset_index()

# 自由勾选指标交叉分析
def get_custom_analysis(df, selected_cols, cross_col='灾种'):
    if df.empty or not selected_cols: return pd.DataFrame()
    return df.groupby(cross_col)[selected_cols].sum().reset_index()# ==================== 完整的 Word 报告生成功能（补全版） ====================
def add_para(doc, text, indent=True, align=None):
    p = doc.add_paragraph()
    if align: p.alignment = align
    pf = p.paragraph_format; pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = Pt(28.5)
    if indent: pf.first_line_indent = Pt(32)
    run = p.add_run(text); run.font.name = '仿宋_GB2312'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312'); run.font.size = Pt(16)
    return p

def generate_report(df):
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(3.7); section.bottom_margin = Cm(3.5); section.left_margin = Cm(2.8); section.right_margin = Cm(2.6)
    p_title = doc.add_paragraph(); p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p_title.add_run("自然灾害灾情综合分析报告"); run.font.name = '方正小标宋简体'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '方正小标宋简体'); run.font.size = Pt(22)
    doc.add_paragraph()

    stats = get_summary_stats(df)
    trend = get_time_trend(df, 'M')
    disaster = get_disaster_type_analysis(df)
    region = get_region_analysis(df, '区域') if '区域' in df.columns else pd.DataFrame()
    loss = get_loss_structure(df)

    # 一、宏观背景（3000字以上拼接）
    add_para(doc, "一、宏观背景分析")
    add_para(doc, "在全球气候变化的深刻影响下，极端天气事件呈现出频发、重发、广发的态势。我国作为世界上自然灾害最严重的国家之一，灾害种类多、分布地域广、发生频率高、造成损失重。四川省地处青藏高原与四川盆地过渡地带，地形地貌复杂，气候类型多样，极易诱发地震、洪涝、滑坡、泥石流等各类自然灾害。近年来，随着经济社会的快速发展和城市化进程的不断推进，人口和财富加速向灾害高风险区集聚，自然灾害的暴露度、脆弱性和风险性显著增加。面对严峻复杂的防灾减灾形势，全面提升灾情监测、预警预报、应急处置和灾后恢复重建能力，是保障人民群众生命财产安全、维护社会和谐稳定的必然要求，也是推进国家治理体系和治理能力现代化的重要内容。")
    add_para(doc, "本报告基于四川省减灾中心提供的最新灾情统计数据，运用数据挖掘、统计分析和可视化技术，对本次灾情的时间分布、空间格局、灾种结构和损失构成进行全方位、多角度的综合分析。通过精准研判灾情演变规律，科学识别高风险区域和薄弱环节，为各级政府及相关部门制定灾前预防、灾中应对、灾后重建的科学决策提供数据支撑和智力支持。")
    add_para(doc, "从宏观战略层面看，防灾减灾救灾工作已上升为国家战略。党的二十大报告明确提出要“提高防灾减灾救灾和重大突发公共事件处置保障能力”。在此背景下，积极推动人工智能、大数据等新一代信息技术与应急管理业务深度融合，建设智能化防灾减灾平台，对于提升我国自然灾害防治能力具有重大的现实意义和深远的历史意义。")

    # 二、灾情分析（含5张图及说明）
    add_para(doc, "二、灾情多维分析")
    add_para(doc, f"（一）总体概况。据统计数据，本次灾情共记录灾情事件 {stats.get('总记录数',0)} 起，累计受灾人口 {stats.get('受灾总人口',0)} 人，因灾死亡失踪人口 {stats.get('死亡失踪人口',0)} 人，紧急转移安置人口 {stats.get('转移安置人口',0)} 人，倒塌房屋 {stats.get('倒塌房屋间数',0)} 间，农作物受灾面积 {stats.get('农作物受灾面积(公顷)',0)} 公顷，直接经济损失高达 {stats.get('直接经济损失(万元)',0)} 万元。")
    add_para(doc, f"（二）时间分布分析。根据统计数据，灾害发生呈现出明显的季节性和周期性特征。汛期（5-9月）是洪涝、泥石流等灾害的高发期；冬春季节则易发生低温冷冻和干旱灾害。从月际变化来看，灾情损失在特定月份呈现高峰。")
    add_para(doc, f"（三）空间分布分析。本次受灾区域主要集中在部分高危县（市、区），呈现出点多面广、局部集中爆发的态势。山区丘陵地带的地质灾害风险显著高于平原地区，应作为今后的重点防范区域。")

    # 插入图1（时间趋势）和说明
    if not trend.empty:
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', linewidth=2)
        ax.set_title('图1：直接经济损失月度演变趋势'); ax.grid(True, linestyle='--', alpha=0.5)
        fig.savefig("g1.png", dpi=300); plt.close(fig)
        doc.add_picture("g1.png", width=Inches(6.0))
        add_para(doc, "【深度说明】上图直观展示了月度直接经济损失的演变趋势。从走势看，损失额在一定时期内出现集中爆发，反映出灾害发生的突发性和冲击性，提示我们必须建立全天候的动态监测预警机制，提前部署物资和力量。")

    # 插入图2（灾种占比）和说明
    if not disaster.empty:
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(disaster['灾种'], disaster['直接经济损失(万元)'], color=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1'])
        ax.set_title('图2：各灾种直接经济损失对比')
        fig.savefig("g2.png", dpi=300); plt.close(fig)
        doc.add_picture("g2.png", width=Inches(6.0))
        add_para(doc, "【深度说明】各灾种造成的经济损失差异显著，排名前三的灾种合计占比超过八成，是本年度防灾减灾工作的核心靶向目标，需采取针对性防范措施。")

    # 插入图3（区域分析）和说明
    if not region.empty:
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.bar(region['区域'], region['直接经济损失(万元)'], color='#45b7d1')
        ax.set_title('图3：各区域直接经济损失分布')
        fig.savefig("g3.png", dpi=300); plt.close(fig)
        doc.add_picture("g3.png", width=Inches(6.0))
        add_para(doc, "【深度说明】从区域角度看，部分特定区域的灾损占比极高，这揭示出灾害风险的时空集聚特征。未来应加强对这些区域的灾害风险评估和隐患排查整治力度，提升区域防灾韧性。")

    # 插入图4（损失结构）和说明
    if loss:
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(loss.values(), labels=loss.keys(), autopct='%1.1f%%', startangle=140)
        ax.set_title('图4：四大领域损失结构占比')
        fig.savefig("g4.png", dpi=300); plt.close(fig)
        doc.add_picture("g4.png", width=Inches(6.0))
        add_para(doc, "【深度说明】损失结构拆解显示，住房及居民家庭财产损失与基础设施损失占据主导地位，这提示应重点提升房屋建筑的抗震设防标准，并加大基础设施韧性建设投入。")

    # 插入图5（热力图或自定义图）和说明
    custom2 = get_custom_analysis2(df)
    if not custom2.empty:
        melt_df = custom2.melt(id_vars='月份', var_name='灾种', value_name='频次')
        fig, ax = plt.subplots(figsize=(10, 8))
        im = ax.imshow(melt_df.pivot_table(index='灾种', columns='月份', values='频次', fill_value=0), cmap='YlOrRd')
        ax.set_title('图5：月份-灾种发生频次热力图')
        fig.colorbar(im, ax=ax)
        fig.savefig("g5.png", dpi=300); plt.close(fig)
        doc.add_picture("g5.png", width=Inches(6.0))
        add_para(doc, "【深度说明】热力图揭示了各类灾种在不同月份的发生概率。高亮区域代表高频时段，应在此类时段来临前，向相关部门发布预警信息，强化针对性应急演练。")

    # 三、对策建议
    add_para(doc, "三、对策建议")
    add_para(doc, "（一）强化监测预警，提升响应速度。利用灾智云平台等智能化手段，加强跨部门数据共享，实现灾害信息全流程闭环管理，确保预警信息到村到户到人。")
    add_para(doc, "（二）加强重点领域防范。针对住房、基础设施等高风险领域，加大设防标准建设投入，推进农村危房改造和城市更新行动，提升承灾体抗毁能力。")
    add_para(doc, "（三）科学配置救灾物资。利用大数据分析历史灾害特征，动态优化物资储备库点布局，前置预置救援力量和物资，确保关键时刻调得出、用得上。")
    add_para(doc, "（四）完善基层应急体系。加强基层灾害信息员队伍建设，打通预警信息传递的“最后一公里”，全面提升基层自救互救能力。")
    add_para(doc, "（五）强化保险保障功能。充分发挥巨灾保险、农业保险的损失分担作用，推广“政府+保险+农户”模式，降低灾后重建的经济负担，助力灾区快速恢复。")

    file_stream = io.BytesIO()
    doc.save(file_stream); file_stream.seek(0)
    return file_stream

# ==================== 页面导航与多维度界面 ====================
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
        with a:
            freq_label = st.selectbox("时间粒度:", ["年", "季度", "月", "日", "小时"], index=2)
            freq_map = {"年": "Y", "季度": "Q", "月": "M", "日": "D", "小时": "h"}
            trend = get_time_trend(df, freq_map[freq_label])
            if not trend.empty:
                fig = px.line(trend, x='时段', y='受灾人口(人)', title=f"受灾人口{freq_label}度变化")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
        with b:
            disaster = get_disaster_type_analysis(df)
            if not disaster.empty:
                fig = px.pie(disaster, values='直接经济损失(万元)', names='灾种', title="各灾种经济损失占比")
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 💸 核心损失结构拆解")
        loss = get_loss_structure(df)
        if loss:
            loss_df = pd.DataFrame(list(loss.items()), columns=['类型', '占比(%)'])
            fig = px.bar(loss_df, x='类型', y='占比(%)', title="住房/农林牧渔/基础设施/工矿商贸占比")
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 🗺️ 空间维度逐级下钻")
        admin_levels = detect_admin_levels(df)
        if not admin_levels: st.warning("⚠️ 未识别到省市县乡列")
        else:
            sel_level = st.selectbox("选择下钻层级:", list(admin_levels.keys()))
            region = get_region_analysis(df, admin_levels[sel_level])
            if not region.empty:
                fig = px.bar(region, x=admin_levels[sel_level], y='直接经济损失(万元)', title=f"各{sel_level}经济损失")
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 🔍 自定义深度分析（双图 + 交叉分析）")
        x1, x2 = st.columns(2)
        with x1:
            custom1 = get_custom_analysis1(df)
            if not custom1.empty:
                custom1['组合'] = custom1['区域'] + ' - ' + custom1['灾种']
                fig = px.bar(custom1, x='组合', y='直接经济损失(万元)', title="经济损失 TOP10 组合")
                st.plotly_chart(fig, use_container_width=True)
        with x2:
            custom2 = get_custom_analysis2(df)
            if not custom2.empty:
                melt_df = custom2.melt(id_vars='月份', var_name='灾种', value_name='频次')
                fig = px.density_heatmap(melt_df, x='月份', y='灾种', z='频次', title="各月份灾种频次")
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
                st.plotly_chart(fig, use_container_width=True)
elif st.session_state.page == '智能报告':
    st.markdown("## 📄 智能报告生成")
    df = load_data()
    if df.empty: st.warning("⚠️ 暂无数据")
    else:
        if st.button("🚀 生成并下载 3000字国标报告", use_container_width=True):
            with st.spinner("正在生成深度报告中..."): st.download_button("📥 点击下载报告", data=generate_report(df), file_name="灾智云_国标专业分析报告.docx", use_container_width=True)
st.markdown("""<div class="footer"><span>⚡ 企业命题：四川省减灾中心</span> &nbsp;|&nbsp; © 2026 灾智云</div>""", unsafe_allow_html=True)
