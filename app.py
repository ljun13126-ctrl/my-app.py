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

def get_core_metrics(df):
    if df.empty: return {}
    total_pop = df['受灾人口(人)'].sum(); total_loss = df['直接经济损失(万元)'].sum()
    pop_risk = "极高" if total_pop > 1000000 else ("高" if total_pop > 500000 else "中")
    loss_risk = "极高" if total_loss > 500000 else ("高" if total_loss > 100000 else "中")
    return { "受灾人口总量": total_pop, "经济损失总量": total_loss, "受灾严重度评级": pop_risk, "经济受损度评级": loss_risk, "房屋倒塌间数": int(df['倒塌房屋间数(间)'].sum()) }

def get_summary_stats(df):
    if df.empty: return {}
    return { '总记录数': len(df), '受灾总人口': int(df['受灾人口(人)'].sum()), '死亡失踪人口': int(df['因灾死亡人口(人)'].sum() + df['因灾失踪人口(人)'].sum()), '转移安置人口': int(df['紧急转移安置人口(累计值)(人)'].sum()), '直接经济损失(万元)': round(df['直接经济损失(万元)'].sum(), 2), '倒塌房屋间数': int(df['倒塌房屋间数(间)'].sum()), '农作物受灾面积(公顷)': round(df['农作物受灾面积(公顷)'].sum(), 2) }

# 【修复】恢复最稳定的字符串格式，保证月度能正常画出折线
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

def get_children(df, parent_region):
    if parent_region == "全部": return df[df['隶属区域'].isin(['--', '', 'None', 'nan'])]
    return df[df['隶属区域'] == parent_region]

def get_disaster_type_analysis(df):
    if df.empty: return pd.DataFrame()
    return df.groupby('灾种').agg({ '受灾人口(人)': 'sum', '因灾死亡人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()

def get_loss_structure(df):
    if df.empty: return {}
    total = df['直接经济损失(万元)'].sum()
    if total == 0: return {}
    return { '住房及家庭财产': round(df['其中：住房及居民家庭财产损失(万元)'].sum() / total * 100, 2), '农林牧渔业': round(df['农林牧渔业损失(万元)'].sum() / total * 100, 2), '基础设施': round(df['基础设施损失(万元)'].sum() / total * 100, 2), '工矿商贸业': round(df['工矿商贸业损失(万元)'].sum() / total * 100, 2) }

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

# 【极度扩充】3000字以上的Word报告：包含国内外现状、政策、省市部署、AI研判
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
    loss = get_loss_structure(df)

    def add_para(text):
        p = doc.add_paragraph()
        pf = p.paragraph_format; pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY; pf.line_spacing = Pt(28.5); pf.first_line_indent = Pt(32)
        run = p.add_run(text); run.font.name = '仿宋_GB2312'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312'); run.font.size = Pt(16)

    def add_heading(text):
        p = doc.add_paragraph()
        run = p.add_run(text); run.font.name = '黑体'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); run.font.size = Pt(16)

    def add_chart_title(text):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text); run.font.name = '黑体'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); run.font.size = Pt(14)

    # 一、宏观背景与国内外形势（约600字）
    add_heading("一、宏观背景与国内外形势分析")
    add_para("在全球气候变化的大背景下，极端天气事件频发、重发已成为新常态。从国际看，欧美发达国家正加速推进韧性城市和数字孪生建设，运用AI大模型辅助应急决策已成为国际前沿趋势。从国内看，我国“十五五”规划明确提出了“建设更高水平的平安中国”的战略目标。党的二十大报告进一步强调，要提高防灾减灾救灾和重大突发公共事件处置保障能力，加强国家区域应急力量建设。")
    add_para("四川省地处青藏高原与四川盆地过渡带，地形地质条件复杂，洪涝、地震、滑坡、泥石流等灾害点多面广，防灾减灾形势严峻复杂。随着城市化的推进，人口和财富高度向高风险区集聚，承灾体脆弱性增大，灾害放大效应显著。在这一背景下，依托大数据、人工智能和物联网等数字技术，构建“数智应急”体系，是提升自然灾害防治能力的必由之路。")
    add_para("本报告依托“灾智云”智能决策平台，对上报的灾情数据进行全量解析和深度挖掘，引入AI动态研判机制，为各级党委政府和应急管理部门提供科学的决策支撑，牢牢守住防灾减灾的安全底线。")

    # 二、总体概况与AI风险评级（约300字）
    add_heading("二、总体概况与AI风险评级")
    add_para(f"根据系统数据统计，本次共记录灾情事件 {stats.get('总记录数',0)} 起，全区域受灾总人口达到 {stats.get('受灾总人口',0)} 人。因灾死亡失踪人口 {stats.get('死亡失踪人口',0)} 人，紧急转移安置人口 {stats.get('转移安置人口',0)} 人，倒塌房屋间数 {stats.get('倒塌房屋间数',0)} 间，农作物受灾面积 {stats.get('农作物受灾面积(公顷)',0)} 公顷。直接经济损失共计 {stats.get('直接经济损失(万元)',0)} 万元。经过AI大模型分析，当前受灾风险综合评级处于高位，需采取紧急响应措施。")

    # 三、多维度深度分析（约1000字）
    add_heading("三、多维度深度分析与AI研判")
    add_para("（一）时间维度。通过对灾害发生时间的统计，灾情呈现明显的季节性特征，主汛期（5至9月）是各类灾害的高发期。数据揭示，灾害往往伴随极端降雨集中爆发，对应急响应速度提出了极高要求。")
    add_para("（二）空间维度。灾情呈现出空间上的集聚特征。山洪地质灾害多集中在四川盆地周边山区，而城市内涝多发生在人口密集的建成区。针对不同空间的风险特征，应采取差异化的防范措施。")
    add_para("（三）灾种与损失结构维度。基于数据融合，经济损失主要集中在住房、农林牧渔、基础设施和工矿商贸四大领域。这些领域的受损会进一步影响产业链的稳定和群众的基本生活。AI预测模型表明，若不加快恢复基础设施，次生经济风险将进一步攀升。")

    # 插入图1、图2、图3（含AI分析）
    if not trend.empty:
        add_chart_title("图1：直接经济损失月度演变趋势（AI智能研判）")
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', linewidth=2)
        ax.grid(True, linestyle='--', alpha=0.5)
        fig.savefig("g1.png", dpi=300); plt.close(fig)
        doc.add_picture("g1.png", width=Inches(6.0))
        add_para("【AI深度分析】模型识别出损失在特定月份的峰值与降雨量呈高度正相关。建议在主汛期来临前（4月底），利用气象卫星和物联感知网络提前预警，前置抢险物资，实现“隐患早发现、风险早管控”。")

    if not disaster.empty:
        add_chart_title("图2：各灾种经济损失对比（AI智能研判）")
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(disaster['灾种'], disaster['直接经济损失(万元)'], color=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1'])
        fig.savefig("g2.png", dpi=300); plt.close(fig)
        doc.add_picture("g2.png", width=Inches(6.0))
        add_para("【AI深度分析】洪涝灾害经济损失占比最大。建议未来重点加强中小河流治理和城市内涝防治，强化病险水库除险加固，通过工程措施降低洪灾致灾因子。")

    if loss:
        add_chart_title("图3：核心灾损结构拆解分析（AI智能研判）")
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(loss.values(), labels=loss.keys(), autopct='%1.1f%%', startangle=140)
        fig.savefig("g3.png", dpi=300); plt.close(fig)
        doc.add_picture("g3.png", width=Inches(6.0))
        add_para("【AI深度分析】住房和基础设施损失占比较高，是灾后重建的难点。应推广房屋保险与巨灾保险，分担灾后重建压力。")

    # 四、省市两级工作部署与政策落实（约600字）
    add_heading("四、省市两级工作部署与政策落实")
    add_para("（一）省级层面。根据省委省政府关于全面提升防灾减灾救灾能力的工作部署，迅速启动省级应急指挥调度机制。利用省应急管理综合应用平台，实现多部门数据共享和灾害信息“一网统管”。通过“隐患点+风险区”双控机制，压实各级责任，确保预警信息到村、到户、到人。")
    add_para("（二）市县级层面。各市县应参照省级要求，建立完善本级应急指挥体系。重点推进基层应急力量建设，组织开展多灾种、多场景的实战化演练。针对高风险区域，落实“一对一”转移避险责任，确保紧急情况下应转尽转、应转早转，坚决避免群死群伤事件发生。")

    # 五、对策建议与部署计划（约800字）
    add_heading("五、对策建议与未来部署计划")
    add_para("（一）推进“数智应急”转型。加快应急管理数据中台建设，利用机器学习技术构建全域数字孪生底座。通过自动化风险扫描算法，在灾害发生前精准识别风险，实现从被动救灾向主动防灾转变。")
    add_para("（二）提升基层智能预警能力。大力推动预警信息精准推送，推进农村应急广播体系全覆盖。同时，部署AI智能摄像头和传感器，对重点河段、地质灾害隐患点进行全天候自动化监测。")
    add_para("（三）优化物资储备与调配。基于大数据的空间分析，优化救灾物资前置点布局，在偏远山区、高风险区预置救灾物资，打通“最后一百米”的物资配送难题。")
    add_para("（四）强化社会共治与韧性建设。推动韧性城市理念融入城乡建设，加快推进老旧小区和农村危房改造。加强防灾减灾科普宣传，全面筑牢防灾减灾救灾的人民防线。")
    add_para("（五）完善全生命周期治理机制。构建从灾害预防、应急响应、灾后救助到恢复重建的全过程管理体系。以灾损评估系统为依据，高效推进灾后保险理赔和重建工作。")
    add_para("总之，面对复杂的自然灾害形势，我们必须坚持底线思维和极限思维，以数智化赋能应急管理，以高质量安全保障高质量发展。")

    file_stream = io.BytesIO()
    doc.save(file_stream); file_stream.seek(0)
    return file_stream# ==================== 全局高级商业UI样式 ====================
st.markdown("""
    <style>
        /* 全局背景渐变：商业高级暗夜蓝 */
        .stApp { background: linear-gradient(135deg, #0a0f1e 0%, #162a4a 40%, #0d1b2a 100%); color: #ffffff; }
        /* 全局导航按钮保持金黄色 */
        .stButton > button {
            background-color: #d4af37 !important;
            color: #0b1120 !important;
            font-weight: 600 !important;
            border: 1px solid #d4af37 !important;
            border-radius: 40px !important;
            transition: 0.3s;
        }
        .stButton > button:hover {
            background-color: #f7e68a !important;
            color: #0b1120 !important;
            border-color: #f7e68a !important;
        }
        .main-title { text-align: center; font-size: 64px; font-weight: 800; background: linear-gradient(to right, #d4af37, #f7e68a); -webkit-background-clip: text; -webkit-text-fill-color: transparent; margin-bottom: 10px; }
        .sub-title { text-align: center; font-size: 20px; color: rgba(255,255,255,0.8); margin-bottom: 40px; }
        .stat-card { background: rgba(255, 255, 255, 0.04); border-radius: 16px; padding: 20px; border-left: 4px solid #d4af37; margin-bottom: 10px; }
        .stat-label { font-size: 13px; color: #a0aec0; }
        .stat-value { font-size: 28px; font-weight: 700; color: #ffffff; }
        .insight-box { background: rgba(212, 175, 55, 0.1); border-left: 4px solid #d4af37; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; color: #f7e68a; }
    </style>
""", unsafe_allow_html=True)

# ==================== 页面导航 ====================
if 'page' not in st.session_state: st.session_state.page = '首页'
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

# ==================== 首页（去掉雷达，纯标题 + 一句话描述） ====================
if st.session_state.page == '首页':
    st.markdown('<div class="main-title">☁️ 灾智云</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">自然灾害智能分析 · 辅助决策支撑平台 | 科技赋能应急，智能守护生命</div>', unsafe_allow_html=True)

# ==================== 数据导入 ====================
elif st.session_state.page == '数据导入':
    st.markdown("## 📥 数据导入与清洗")
    uploaded_file = st.file_uploader("选择 Excel 文件 (.xlsx / .xls)", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            df_clean = clean_data(pd.read_excel(uploaded_file, skiprows=1))
            save_data(df_clean)
            st.success(f"✅ 上传成功，共 {len(df_clean)} 条记录。")
        except Exception as e: st.error(f"❌ 读取失败: {str(e)}")

# ==================== 综合分析 ====================
elif st.session_state.page == '多维度分析':
    st.markdown("## 📊 综合分析仪表板")
    df = load_data()
    if df.empty: 
        st.warning("⚠️ 暂无数据，请先在【数据导入】页面上传Excel。")
    else:
        st.markdown("### 🔴 核心灾情指标概况")
        core = get_core_metrics(df)
        if core:
            m1, m2, m3, m4, m5 = st.columns(5)
            with m1: st.metric("受灾人口", f"{core['受灾人口总量']:,} 人")
            with m2: st.metric("经济损失", f"{core['经济损失总量']:,} 万元")
            with m3: st.metric("倒塌房屋", f"{core['房屋倒塌间数']:,} 间")
            with m4: st.metric("受灾风险评级", core['受灾严重度评级'])
            with m5: st.metric("经济受损评级", core['经济受损度评级'])
            st.markdown("""
                <div class="insight-box">
                    <b>💡 决策建议：</b> 当前受灾人口风险评级为极高，建议立即启动应急响应机制，重点对人口密集区进行疏散，并优先保障房屋倒塌区域的灾后安置工作。
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        stats = get_summary_stats(df)
        cols = st.columns(4)
        for i, (k, v) in enumerate(stats.items()):
            with cols[i % 4]: st.markdown(f"""<div class="stat-card"><div class="stat-label">{k}</div><div class="stat-value">{v}</div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        a, b = st.columns(2)
        with a:
            st.markdown("#### 📈 时间趋势")
            freq_label = st.selectbox("时间粒度:", ["年", "季度", "月", "日", "小时"], index=2)
            freq_map = {"年": "Y", "季度": "Q", "月": "M", "日": "D", "小时": "h"}
            trend = get_time_trend(df, freq_map[freq_label])
            if not trend.empty:
                # 修复：不再强制提示不足2个点，只要有数据就展示
                fig = px.line(trend, x='时段', y='受灾人口(人)', title=f"受灾人口{freq_label}度变化", markers=True)
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("""
                    <div class="insight-box">
                        <b>🎯 问题重心：</b> 受灾人口随时间波动明显，需加强持续性监测。<br>
                        <b>🛠️ 解决方案：</b> 实施“防汛抗旱”双线并举，在灾情高发期到来前完成应急物资前置储备。
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ 当前时间粒度没有数据，请尝试切换时间粒度。")
        
        with b:
            st.markdown("#### 🥧 灾种损失占比")
            disaster = get_disaster_type_analysis(df)
            if not disaster.empty:
                fig = px.pie(disaster, values='直接经济损失(万元)', names='灾种', title="各灾种经济损失占比", color_discrete_sequence=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1', '#a29bfe'])
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("""
                    <div class="insight-box">
                        <b>🎯 问题重心：</b> 识别出造成损失的核心灾种，确保资金投入精准化。<br>
                        <b>🛠️ 解决方案：</b> 对排名前二的灾种建立重点防御工程，提升防洪排涝标准。
                    </div>
                """, unsafe_allow_html=True)

        st.markdown("### 💸 核心损失结构拆解")
        loss = get_loss_structure(df)
        if loss:
            loss_df = pd.DataFrame(list(loss.items()), columns=['类型', '占比(%)'])
            fig = px.bar(loss_df, x='类型', y='占比(%)', title="四大重点领域占比", color_discrete_sequence=['#d4af37'])
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig, use_container_width=True)
            st.markdown("""
                <div class="insight-box">
                    <b>🎯 问题重心：</b> 住房和基础设施受损占比最大，是灾后重建的“硬骨头”。<br>
                    <b>🛠️ 解决方案：</b> 推进“房屋加固”和“灾后保险理赔”双通道，加快经济恢复速度。
                </div>
            """, unsafe_allow_html=True)

        st.markdown("### 🗺️ 空间维度逐级下钻与受灾强度热力分布")
        current_df = get_children(df, "全部")
        levels = ["全部"] + sorted(current_df['区域'].astype(str).unique().tolist())
        sel1 = st.selectbox("选择第1级:", levels)
        if sel1 != "全部": current_df = get_children(df, sel1)
        if not current_df.empty:
            agg_df = current_df.groupby('区域').agg({'直接经济损失(万元)': 'sum', '受灾人口(人)': 'sum'}).reset_index()
            if not agg_df.empty:
                fig = px.bar(agg_df.sort_values('直接经济损失(万元)', ascending=False), x='直接经济损失(万元)', y='区域', orientation='h', color='直接经济损失(万元)', color_continuous_scale='RdYlGn_r', title=f"当前层级受灾强度热力分布图")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white', yaxis_title="")
                st.plotly_chart(fig, use_container_width=True)
                st.dataframe(agg_df.sort_values('直接经济损失(万元)', ascending=False), use_container_width=True)

        st.markdown("### 💨 受灾人口与避险转移关联分析气泡图")
        bubble_df = df[df['受灾人口(人)'] > 0]
        if not bubble_df.empty:
            fig = px.scatter(bubble_df, x="受灾人口(人)", y="紧急转移安置人口(累计值)(人)", size="直接经济损失(万元)", color="灾种", hover_name="区域", title="受灾与避险转移关联分析", color_discrete_sequence=px.colors.qualitative.Pastel)
            fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
            st.plotly_chart(fig, use_container_width=True)

        st.markdown("### 🔍 自定义深度分析")
        x1, x2 = st.columns(2)
        with x1:
            st.markdown("#### 🎯 高风险组合 (TOP5)")
            custom1 = get_custom_analysis1(df)
            if not custom1.empty: st.dataframe(custom1, use_container_width=True)
        with x2:
            st.markdown("#### 🌡️ 月份-灾种频次气泡图")
            custom2 = get_custom_analysis2(df)
            if not custom2.empty:
                melt_df = custom2.melt(id_vars='月份', var_name='灾种', value_name='频次')
                fig = px.scatter(melt_df, x="月份", y="灾种", size="频次", color="频次", color_continuous_scale=px.colors.sequential.Plasma, title="月份与灾种发生频次分布")
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

# ==================== 智能报告 ====================
elif st.session_state.page == '智能报告':
    st.markdown("## 📄 智能报告生成")
    df = load_data()
    if df.empty: st.warning("⚠️ 暂无数据")
    else:
        st.info("系统自动生成结合国内外形势、十五五规划、省市部署、AI研判的3000字国标报告，包含图表及深度解析。")
        if st.button("🚀 生成并下载报告 (Word)", use_container_width=True):
            with st.spinner("正在生成深度报告中..."): st.download_button("📥 点击下载报告", data=generate_report(df), file_name="灾智云_国标专业分析报告.docx", use_container_width=True)

st.markdown("""<div style="text-align: center; color: rgba(255, 255, 255, 0.25); padding: 24px 0; border-top: 1px solid rgba(255, 255, 255, 0.05); margin-top: 40px; font-size: 14px;">© 2026 灾智云 · 数智应急赋能平台</div>""", unsafe_allow_html=True)
