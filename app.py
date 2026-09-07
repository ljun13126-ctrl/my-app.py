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

def get_time_trend(df, freq='M'):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    if freq == 'Y': df_t['时段'] = df_t['时间'].dt.to_period('Y').dt.to_timestamp()
    elif freq == 'Q': df_t['时段'] = df_t['时间'].dt.to_period('Q').dt.to_timestamp()
    elif freq == 'M': df_t['时段'] = df_t['时间'].dt.to_period('M').dt.to_timestamp()
    elif freq == 'D': df_t['时段'] = df_t['时间'].dt.normalize()
    elif freq == 'h': df_t['时段'] = df_t['时间'].dt.floor('h')
    grouped = df_t.groupby('时段').agg({ '受灾人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()
    return grouped.sort_values('时段')

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
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text); run.font.name = '黑体'; run._element.rPr.rFonts.set(qn('w:eastAsia'), '黑体'); run.font.size = Pt(14)

    add_heading("一、宏观背景与战略意义")
    add_para("当前，我国正处于“十五五”规划谋篇布局的关键时期，也是全面推进国家治理体系和治理能力现代化的攻坚阶段。党的二十大报告明确提出要提高防灾减灾救灾和重大突发公共事件处置保障能力。在此背景下，应急管理体系的数智化转型已成为提升国家治理效能的必然要求。四川省地处青藏高原与四川盆地过渡带，地形复杂，灾害频发，灾害风险交织叠加。深入剖析灾情数据，利用大数据、人工智能等现代化手段辅助决策，是践行“人民至上、生命至上”理念的具体实践。")
    add_para("本报告利用“灾智云”智能决策平台，对近期上报的灾情数据进行全维度解析，旨在摸清灾害底数，揭示损失规律，为各级政府和应急管理部门开展精准救灾、科学决策提供强有力的数据支撑。")

    add_heading("二、灾情总体概况与风险评级")
    add_para(f"根据系统数据统计，本次共记录灾情事件 {stats.get('总记录数',0)} 起，全区域受灾总人口达到 {stats.get('受灾总人口',0)} 人。因灾死亡失踪人口 {stats.get('死亡失踪人口',0)} 人，紧急转移安置人口 {stats.get('转移安置人口',0)} 人，倒塌房屋间数 {stats.get('倒塌房屋间数',0)} 间，农作物受灾面积 {stats.get('农作物受灾面积(公顷)',0)} 公顷。直接经济损失共计 {stats.get('直接经济损失(万元)',0)} 万元。基于风险模型综合评估，当前四川省整体受灾风险处于高位。")

    add_heading("三、多维度深度分析与图表解析")
    
    if not trend.empty:
        add_chart_title("图1：直接经济损失月度演变趋势")
        fig, ax = plt.subplots(figsize=(10, 5))
        ax.plot(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', linewidth=2)
        ax.set_title('直接经济损失月度演变趋势'); ax.grid(True, linestyle='--', alpha=0.5)
        fig.savefig("g1.png", dpi=300); plt.close(fig)
        doc.add_picture("g1.png", width=Inches(6.0))
        add_para("【深度解析与决策建议】从月度变化曲线看，经济损失在特定时段出现明显波峰，这与汛期集中降水高度吻合。建议：在主汛期前（4月底）前置抢险救援力量和物资；利用气象卫星和物联感知网络提升预警提前量，实现“隐患早发现、风险早管控”。")

    if not disaster.empty:
        add_chart_title("图2：各灾种经济损失对比")
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(disaster['灾种'], disaster['直接经济损失(万元)'], color=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1'])
        fig.savefig("g2.png", dpi=300); plt.close(fig)
        doc.add_picture("g2.png", width=Inches(6.0))
        add_para("【深度解析与决策建议】洪涝等灾害经济损失占据绝对主导地位。建议：重点加强中小河流治理和城市内涝防治，强化病险水库除险加固，从源头上减少洪水灾害对重点区域的侵袭。")

    if loss:
        add_chart_title("图3：核心灾损结构拆解分析")
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(loss.values(), labels=loss.keys(), autopct='%1.1f%%', startangle=140)
        fig.savefig("g3.png", dpi=300); plt.close(fig)
        doc.add_picture("g3.png", width=Inches(6.0))
        add_para("【深度解析与决策建议】损失结构中，住房和基础设施占比最高。建议：加速推进农村危房改造和高标准农田建设，推广巨灾保险制度，有效分担灾后重建经济压力。")

    add_heading("四、结合“十五五”规划与数智应急的对策建议")
    add_para("（一）推进应急管理数智化转型。全面贯彻落实“十五五”规划关于数字中国建设的部署要求，加快应急管理数据中台建设，打通跨部门数据壁垒，实现灾情信息“一网统管”。")
    add_para("（二）提升灾害监测预警智能化水平。以“数智应急”为抓手，利用AI大模型预测灾害演化趋势，构建数字孪生流域和城市，实现从被动救灾向主动防灾转变。")
    add_para("（三）优化基层应急物资储备体系。基于大数据空间分析，优化救灾物资前置点布局，推动应急物资储备库向偏远山区、高风险区延伸，解决“最后一百米”的物资配送难题。")
    add_para("（四）强化宣传教育与应急演练。通过多形式常态化开展防灾减灾宣传，提升全民灾害防范意识和自救互救能力，筑牢防灾减灾救灾的人民防线。")
    add_para("（五）构建全生命周期灾害治理体系。从灾害预防、应急响应、灾后救助到恢复重建，实施全过程精细化管理。引入灾损评估系统，为灾后科学重建和保险理赔提供客观依据。")
    add_para("综上所述，面对复杂的自然灾害形势，必须坚持以大概率思维应对小概率事件，以数智化赋能应急管理，以高质量安全保障高质量发展。")

    file_stream = io.BytesIO()
    doc.save(file_stream); file_stream.seek(0)
    return file_stream# ==================== 页面导航 ====================
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

# ==================== 首页 ====================
if st.session_state.page == '首页':
    # 引入体现“数智应急”的科技感视觉效果（雷达扫描、AI、数据流等）
    st.markdown("""
        <style>
            .hero-banner {
                background: linear-gradient(135deg, rgba(11,17,32,0.9) 0%, rgba(22,42,74,0.9) 100%);
                border: 1px solid rgba(212,175,55,0.3);
                border-radius: 20px;
                padding: 40px;
                margin-bottom: 20px;
                text-align: center;
            }
            .hero-title {
                font-size: 56px;
                font-weight: 800;
                background: -webkit-linear-gradient(#d4af37, #f7e68a);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin-bottom: 10px;
            }
            .hero-sub {
                font-size: 18px;
                color: rgba(255,255,255,0.8);
                margin-bottom: 30px;
            }
            .module-card {
                display: inline-block;
                width: 22%;
                margin: 0 1%;
                padding: 20px;
                border: 1px solid rgba(212,175,55,0.2);
                border-radius: 12px;
                background: rgba(255,255,255,0.03);
                text-align: center;
                transition: 0.3s;
            }
            .module-card:hover {
                transform: translateY(-5px);
                border-color: #d4af37;
                background: rgba(212,175,55,0.1);
            }
            .radar {
                margin: 0 auto 20px auto;
                width: 200px;
                height: 200px;
                border: 2px solid rgba(78,205,196,0.5);
                border-radius: 50%;
                box-shadow: 0 0 40px rgba(78,205,196,0.3);
                position: relative;
                animation: pulse 2s infinite;
                display: flex;
                align-items: center;
                justify-content: center;
                font-size: 80px;
            }
            @keyframes pulse {
                0% { transform: scale(0.95); opacity: 0.8; }
                50% { transform: scale(1.05); opacity: 1; }
                100% { transform: scale(0.95); opacity: 0.8; }
            }
        </style>
        
        <div class="hero-banner">
            <div class="hero-title">☁️ 灾智云 · 数智应急引擎</div>
            <div class="hero-sub">科技赋能应急 · 智能守护生命</div>
            
            <div class="radar">
                🛰️
            </div>
            
            <div>
                <div class="module-card">
                    <div style="font-size: 40px; color: #d4af37;">🤖</div>
                    <div style="font-size: 16px; margin-top: 10px; color: #fff;">AI智能研判</div>
                    <div style="font-size: 12px; margin-top: 5px; color: rgba(255,255,255,0.5);">大模型辅助决策</div>
                </div>
                <div class="module-card">
                    <div style="font-size: 40px; color: #4ecdc4;">📡</div>
                    <div style="font-size: 16px; margin-top: 10px; color: #fff;">实时数据监测</div>
                    <div style="font-size: 12px; margin-top: 5px; color: rgba(255,255,255,0.5);">物联感知互联</div>
                </div>
                <div class="module-card">
                    <div style="font-size: 40px; color: #ff6b6b;">🚨</div>
                    <div style="font-size: 16px; margin-top: 10px; color: #fff;">风险预警引擎</div>
                    <div style="font-size: 12px; margin-top: 5px; color: rgba(255,255,255,0.5);">隐患早发现早管控</div>
                </div>
                <div class="module-card">
                    <div style="font-size: 40px; color: #f7e68a;">🧬</div>
                    <div style="font-size: 16px; margin-top: 10px; color: #fff;">多维智能分析</div>
                    <div style="font-size: 12px; margin-top: 5px; color: rgba(255,255,255,0.5);">多维度数据穿透</div>
                </div>
            </div>
        </div>
    """, unsafe_allow_html=True)

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
        # 核心指标概况
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
                <div style="background: rgba(212, 175, 55, 0.1); border-left: 4px solid #d4af37; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; color: #f7e68a;">
                    <b>💡 决策建议：</b> 当前受灾人口风险评级为极高，建议立即启动应急响应机制，重点对人口密集区进行疏散，并优先保障房屋倒塌区域的灾后安置工作。
                </div>
            """, unsafe_allow_html=True)
        
        st.markdown("---")
        stats = get_summary_stats(df)
        cols = st.columns(4)
        for i, (k, v) in enumerate(stats.items()):
            with cols[i % 4]: st.markdown(f"""<div style="background: rgba(255, 255, 255, 0.04); border-radius: 16px; padding: 20px 24px; border-left: 4px solid #d4af37; margin-bottom: 10px;"><div style="font-size: 13px; color: #a0aec0;">{k}</div><div style="font-size: 28px; font-weight: 700; color: #fff;">{v}</div></div>""", unsafe_allow_html=True)
        
        st.markdown("---")
        a, b = st.columns(2)
        
        # 时间趋势分析
        with a:
            st.markdown("#### 📈 时间趋势")
            freq_label = st.selectbox("时间粒度:", ["年", "季度", "月", "日", "小时"], index=2)
            freq_map = {"年": "Y", "季度": "Q", "月": "M", "日": "D", "小时": "h"}
            trend = get_time_trend(df, freq_map[freq_label])
            if not trend.empty:
                if len(trend) < 2:
                    st.warning("⚠️ 当前数据点不足2个，无法画出折线，请尝试切换更细粒度或补充数据！")
                else:
                    fig = px.line(trend, x='时段', y='受灾人口(人)', title=f"受灾人口{freq_label}度变化", markers=True)
                    fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                    st.plotly_chart(fig, use_container_width=True)
                    st.markdown("""
                        <div style="background: rgba(78, 205, 196, 0.1); border-left: 4px solid #4ecdc4; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; color: #4ecdc4;">
                            <b>🎯 问题重心：</b> 受灾人口随时间波动明显，需加强持续性监测。<br>
                            <b>🛠️ 解决方案：</b> 实施“防汛抗旱”双线并举，在灾情高发期到来前完成应急物资前置储备。
                        </div>
                    """, unsafe_allow_html=True)
        
        # 灾种占比
        with b:
            st.markdown("#### 🥧 灾种损失占比")
            disaster = get_disaster_type_analysis(df)
            if not disaster.empty:
                fig = px.pie(disaster, values='直接经济损失(万元)', names='灾种', title="各灾种经济损失占比", color_discrete_sequence=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1', '#a29bfe'])
                fig.update_layout(paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("""
                    <div style="background: rgba(255, 107, 107, 0.1); border-left: 4px solid #ff6b6b; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; color: #ff6b6b;">
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
                <div style="background: rgba(212, 175, 55, 0.1); border-left: 4px solid #d4af37; padding: 10px 15px; border-radius: 8px; margin-bottom: 20px; color: #f7e68a;">
                    <b>🎯 问题重心：</b> 住房和基础设施受损占比最大，是灾后重建的“硬骨头”。<br>
                    <b>🛠️ 解决方案：</b> 推进“房屋加固”和“灾后保险理赔”双通道，加快经济恢复速度。
                </div>
            """, unsafe_allow_html=True)

        # 空间维度下钻
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

        # 气泡图
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
            if not custom1.empty:
                st.dataframe(custom1, use_container_width=True)
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
        st.info("系统自动生成结合“十五五”规划与数智应急要求的3000字国标报告，包含图表及深度解析。")
        if st.button("🚀 生成并下载报告 (Word)", use_container_width=True):
            with st.spinner("正在生成深度报告中..."): 
                st.download_button("📥 点击下载报告", data=generate_report(df), file_name="灾智云_国标专业分析报告.docx", use_container_width=True)

st.markdown("""<div style="text-align: center; color: rgba(255, 255, 255, 0.25); padding: 24px 0; border-top: 1px solid rgba(255, 255, 255, 0.05); margin-top: 40px; font-size: 14px;">© 2026 灾智云 · 数智应急赋能平台</div>""", unsafe_allow_html=True)
