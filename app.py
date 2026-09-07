import streamlit as st
import pandas as pd
import sqlite3
import os
import io
from docx import Document
import plotly.express as px
import plotly.graph_objects as go

# ==================== 页面配置 ====================
st.set_page_config(page_title="灾智云 · 智能决策平台", layout="wide", page_icon="☁️")

# ===== 自定义 CSS：完全隐藏侧边栏 + 高级商业风格 =====
st.markdown("""
    <style>
        /* 完全隐藏侧边栏 */
        .css-1d391kg { display: none !important; }
        section[data-testid="stSidebar"] { display: none !important; }
        .st-emotion-cache-1r6slb0 { display: none !important; }
        [data-testid="stSidebar"] { display: none !important; }
        
        /* 全局背景 */
        .stApp {
            background: linear-gradient(135deg, #0a0f1e 0%, #162a4a 40%, #0d1b2a 100%);
            color: #ffffff;
        }
        
        /* 主内容区调整 */
        .main > div {
            padding: 1rem 3rem !important;
        }
        
        /* ===== 自定义顶部导航 ===== */
        .nav-container {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 16px 40px;
            background: rgba(11, 17, 32, 0.7);
            backdrop-filter: blur(16px);
            border-bottom: 1px solid rgba(212, 175, 55, 0.15);
            border-radius: 0 0 20px 20px;
            margin-bottom: 20px;
        }
        .nav-brand {
            font-size: 28px;
            font-weight: 700;
            color: #d4af37;
            text-decoration: none;
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .nav-links {
            display: flex;
            gap: 32px;
            list-style: none;
            margin: 0;
            padding: 0;
        }
        .nav-links li a {
            color: rgba(255, 255, 255, 0.65);
            text-decoration: none;
            font-size: 16px;
            font-weight: 500;
            padding: 8px 16px;
            border-radius: 8px;
            transition: 0.3s;
            cursor: pointer;
        }
        .nav-links li a:hover {
            color: #ffffff;
            background: rgba(212, 175, 55, 0.1);
        }
        .nav-links li a.active {
            color: #d4af37;
            background: rgba(212, 175, 55, 0.15);
        }
        
        /* ===== 标题 ===== */
        .main-title {
            font-size: 72px;
            font-weight: 700;
            background: linear-gradient(to right, #d4af37, #f7e68a);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            text-align: center;
            padding-top: 10px;
        }
        .sub-title {
            text-align: center;
            font-size: 24px;
            color: rgba(255, 255, 255, 0.8);
            margin-bottom: 10px;
        }
        
        /* ===== 统计卡片 ===== */
        .stat-card {
            background: rgba(255, 255, 255, 0.04);
            border-radius: 16px;
            padding: 20px 24px;
            border-left: 4px solid #d4af37;
            backdrop-filter: blur(8px);
            border: 1px solid rgba(255, 255, 255, 0.05);
        }
        .stat-label {
            font-size: 13px;
            color: #a0aec0;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .stat-value {
            font-size: 28px;
            font-weight: 700;
            color: #ffffff;
        }
        
        /* ===== 首页步骤卡片 ===== */
        .step-card {
            background: rgba(255, 255, 255, 0.03);
            padding: 24px 30px;
            border-radius: 16px;
            border: 1px solid rgba(212, 175, 55, 0.12);
            text-align: center;
            transition: 0.3s;
            flex: 1;
            min-width: 180px;
        }
        .step-card:hover {
            background: rgba(212, 175, 55, 0.08);
            border-color: #d4af37;
            transform: translateY(-4px);
        }
        .step-number {
            font-size: 32px;
            font-weight: 700;
            color: #d4af37;
            display: block;
            margin-bottom: 8px;
        }
        .step-label {
            color: rgba(255, 255, 255, 0.8);
            font-size: 16px;
            font-weight: 500;
        }
        .step-desc {
            color: rgba(255, 255, 255, 0.4);
            font-size: 13px;
            margin-top: 4px;
        }
        
        /* ===== 页脚 ===== */
        .footer {
            text-align: center;
            color: rgba(255, 255, 255, 0.25);
            padding: 24px 0;
            border-top: 1px solid rgba(255, 255, 255, 0.05);
            margin-top: 40px;
            font-size: 14px;
        }
        .footer span { color: #d4af37; }
        
        /* ===== 按钮 ===== */
        .stButton > button {
            background: #d4af37 !important;
            color: #0b1120 !important;
            font-weight: 600 !important;
            border-radius: 40px !important;
            border: none !important;
            padding: 0.5rem 2.5rem !important;
            transition: 0.3s;
        }
        .stButton > button:hover {
            background: #f7e68a !important;
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(212, 175, 55, 0.35);
        }
        
        /* ===== 表格 ===== */
        .dataframe {
            background: rgba(255, 255, 255, 0.03) !important;
            border-radius: 12px !important;
            color: #fff !important;
        }
        .dataframe thead th {
            background: #1a2a4a !important;
            color: #d4af37 !important;
        }
        
        /* ===== 文件上传 ===== */
        .stFileUploader > div {
            border: 2px dashed rgba(212, 175, 55, 0.3) !important;
            border-radius: 16px !important;
            background: rgba(212, 175, 55, 0.03) !important;
        }
        
        /* ===== 图表容器 ===== */
        .chart-container {
            background: rgba(255, 255, 255, 0.02);
            border-radius: 16px;
            padding: 16px;
            border: 1px solid rgba(255, 255, 255, 0.04);
        }
        
        /* 响应式 */
        @media (max-width: 768px) {
            .main > div { padding: 1rem !important; }
            .nav-container { flex-direction: column; gap: 12px; padding: 12px 20px; }
            .nav-links { flex-wrap: wrap; justify-content: center; gap: 12px; }
            .main-title { font-size: 40px; }
            .sub-title { font-size: 18px; }
            .stat-value { font-size: 22px; }
        }
    </style>
""", unsafe_allow_html=True)

# ==================== 数据库初始化 ====================
DB_PATH = "data/uploaded_data.db"
os.makedirs("data", exist_ok=True)

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS disaster_data (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            region TEXT,
            disaster_type TEXT,
            parent_region TEXT,
            disaster_time TEXT,
            affected_population INTEGER,
            death_population INTEGER,
            missing_population INTEGER,
            emergency_evacuation INTEGER,
            emergency_relocation INTEGER,
            emergency_life_aid INTEGER,
            collapsed_houses INTEGER,
            collapsed_households INTEGER,
            severe_damaged_houses INTEGER,
            severe_damaged_households INTEGER,
            moderate_damaged_houses INTEGER,
            moderate_damaged_households INTEGER,
            crop_area_affected REAL,
            crop_area_no_harvest REAL,
            direct_economic_loss REAL,
            housing_loss REAL,
            agri_loss REAL,
            industry_loss REAL,
            infrastructure_loss REAL,
            public_service_loss REAL,
            other_loss REAL
        )
    ''')
    conn.commit()
    conn.close()
init_db()

def load_data():
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql_query("SELECT * FROM disaster_data", conn)
    conn.close()
    return df

def save_data(df):
    conn = sqlite3.connect(DB_PATH)
    df.to_sql('disaster_data', conn, if_exists='replace', index=False)
    conn.close()

# ==================== 数据清洗 ====================
def clean_data(df):
    df = df.fillna({
        '受灾人口(人)': 0, '因灾死亡人口(人)': 0, '因灾失踪人口(人)': 0,
        '紧急避险转移人口(人)': 0, '紧急转移安置人口(累计值)(人)': 0,
        '需紧急生活救助人口(累计值)(人)': 0, '倒塌房屋间数(间)': 0,
        '倒塌住房户数(户)': 0, '严重损坏房屋间数(间)': 0,
        '严重损坏住房户数(户)': 0, '一般损坏房屋间数(间)': 0,
        '一般损坏住房户数(户)': 0, '农作物受灾面积(公顷)': 0.0,
        '农作物绝收面积(公顷)': 0.0, '直接经济损失(万元)': 0.0,
        '其中：住房及居民家庭财产损失(万元)': 0.0, '农林牧渔业损失(万元)': 0.0,
        '工矿商贸业损失(万元)': 0.0, '基础设施损失(万元)': 0.0,
        '公共服务损失(万元)': 0.0, '其他损失(万元)': 0.0
    })
    num_cols = df.select_dtypes(include=['number']).columns
    for col in num_cols:
        df[col] = df[col].apply(lambda x: max(x, 0) if isinstance(x, (int, float)) else x)
    if '灾害发生时间' in df.columns:
        df['灾害发生时间'] = pd.to_datetime(df['灾害发生时间'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')
    for col in ['区域', '隶属区域']:
        if col in df.columns:
            df[col] = df[col].fillna('--')
    return df

# ==================== 分析函数 ====================
def get_summary_stats(df):
    if df.empty:
        return {}
    return {
        '总记录数': len(df),
        '受灾总人口': int(df['受灾人口(人)'].sum()),
        '死亡失踪人口': int(df['因灾死亡人口(人)'].sum() + df['因灾失踪人口(人)'].sum()),
        '转移安置人口': int(df['紧急转移安置人口(累计值)(人)'].sum()),
        '直接经济损失(万元)': round(df['直接经济损失(万元)'].sum(), 2),
        '倒塌房屋间数': int(df['倒塌房屋间数(间)'].sum()),
        '农作物受灾面积(公顷)': round(df['农作物受灾面积(公顷)'].sum(), 2)
    }

def get_time_trend(df, freq='M'):
    if df.empty or '灾害发生时间' not in df.columns:
        return pd.DataFrame()
    df_t = df.copy()
    df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce')
    df_t = df_t.dropna(subset=['时间'])
    if freq == 'M':
        df_t['时段'] = df_t['时间'].dt.to_period('M').astype(str)
    elif freq == 'D':
        df_t['时段'] = df_t['时间'].dt.date.astype(str)
    else:
        df_t['时段'] = df_t['时间'].dt.year.astype(str)
    grouped = df_t.groupby('时段').agg({
        '受灾人口(人)': 'sum',
        '直接经济损失(万元)': 'sum',
        '倒塌房屋间数(间)': 'sum'
    }).reset_index()
    return grouped

def get_region_analysis(df):
    if df.empty:
        return pd.DataFrame()
    grouped = df.groupby('区域').agg({
        '受灾人口(人)': 'sum',
        '因灾死亡人口(人)': 'sum',
        '直接经济损失(万元)': 'sum',
        '农作物受灾面积(公顷)': 'sum'
    }).reset_index()
    return grouped

def get_disaster_type_analysis(df):
    if df.empty:
        return pd.DataFrame()
    grouped = df.groupby('灾种').agg({
        '受灾人口(人)': 'sum',
        '因灾死亡人口(人)': 'sum',
        '直接经济损失(万元)': 'sum',
        '倒塌房屋间数(间)': 'sum'
    }).reset_index()
    return grouped

def get_loss_structure(df):
    if df.empty:
        return {}
    total = df['直接经济损失(万元)'].sum()
    if total == 0:
        return {}
    losses = {
        '住房及家庭财产': df['其中：住房及居民家庭财产损失(万元)'].sum(),
        '农林牧渔业': df['农林牧渔业损失(万元)'].sum(),
        '工矿商贸业': df['工矿商贸业损失(万元)'].sum(),
        '基础设施': df['基础设施损失(万元)'].sum(),
        '公共服务': df['公共服务损失(万元)'].sum(),
        '其他': df['其他损失(万元)'].sum()
    }
    return {k: round(v / total * 100, 2) for k, v in losses.items()}

def get_custom_analysis1(df):
    if df.empty:
        return pd.DataFrame()
    cross = df.groupby(['区域', '灾种']).agg({'直接经济损失(万元)': 'sum', '受灾人口(人)': 'sum'}).reset_index()
    return cross.sort_values('直接经济损失(万元)', ascending=False).head(10)

def get_custom_analysis2(df):
    if df.empty or '灾害发生时间' not in df.columns:
        return pd.DataFrame()
    df_t = df.copy()
    df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce')
    df_t = df_t.dropna(subset=['时间'])
    df_t['月份'] = df_t['时间'].dt.month
    pivot = df_t.pivot_table(index='月份', columns='灾种', aggfunc='size', fill_value=0).reset_index()
    return pivot

# ==================== 生成 Word 报告 ====================
def generate_report(df):
    doc = Document()
    doc.add_heading('自然灾害灾情综合分析报告', 0)
    doc.add_heading('一、总体概况', level=1)
    stats = get_summary_stats(df)
    p = doc.add_paragraph()
    for k, v in stats.items():
        p.add_run(f'{k}: {v}  ').bold = True
    doc.add_heading('二、分灾种分析', level=1)
    dis_data = get_disaster_type_analysis(df)
    if not dis_data.empty:
        table = doc.add_table(rows=1, cols=5)
        hdr = table.rows[0].cells
        hdr[0].text = '灾种'; hdr[1].text = '受灾人口'; hdr[2].text = '死亡人口'; hdr[3].text = '经济损失(万元)'; hdr[4].text = '倒塌房屋(间)'
        for _, row in dis_data.iterrows():
            cells = table.add_row().cells
            cells[0].text = str(row['灾种'])
            cells[1].text = str(row['受灾人口(人)'])
            cells[2].text = str(row['因灾死亡人口(人)'])
            cells[3].text = str(round(row['直接经济损失(万元)'], 2))
            cells[4].text = str(row['倒塌房屋间数(间)'])
    doc.add_heading('三、损失结构', level=1)
    loss = get_loss_structure(df)
    for k, v in loss.items():
        doc.add_paragraph(f'{k}: {v}%')
    doc.add_heading('四、结论与建议', level=1)
    doc.add_paragraph('基于分析结果，建议加强高风险区域和主要灾种的监测预警，完善应急物资储备，提升基层响应能力。')
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

# ==================== 页面导航 ====================
if 'page' not in st.session_state:
    st.session_state.page = '首页'

# ===== 自定义导航 HTML（仅装饰，实际点击由下方按钮控制） =====
st.markdown(f"""
    <nav class="nav-container">
        <div class="nav-brand">☁️ 灾智云</div>
        <ul class="nav-links">
            <li><a class="{'active' if st.session_state.page == '首页' else ''}">首页</a></li>
            <li><a class="{'active' if st.session_state.page == '数据导入' else ''}">数据导入</a></li>
            <li><a class="{'active' if st.session_state.page == '多维度分析' else ''}">多维度分析</a></li>
            <li><a class="{'active' if st.session_state.page == '智能报告' else ''}">智能报告</a></li>
        </ul>
    </nav>
""", unsafe_allow_html=True)

# ===== 页面切换按钮（使用列布局） =====
def set_page(page_name):
    st.session_state.page = page_name

col1, col2, col3, col4 = st.columns(4)
with col1:
    if st.button("🏠 首页", key="nav_home", use_container_width=True):
        set_page('首页')
with col2:
    if st.button("📥 数据导入", key="nav_upload", use_container_width=True):
        set_page('数据导入')
with col3:
    if st.button("📊 多维度分析", key="nav_analysis", use_container_width=True):
        set_page('多维度分析')
with col4:
    if st.button("📄 智能报告", key="nav_report", use_container_width=True):
        set_page('智能报告')

st.markdown("---")

# ==================== 页面内容 ====================

# ---------- 首页 ----------
if st.session_state.page == '首页':
    st.markdown('<div class="main-title">灾智云</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">自然灾害智能分析 · 辅助决策支撑平台</div>', unsafe_allow_html=True)
    
    st.markdown("""
        <div style="text-align:center; padding:10px 0 30px 0;">
            <p style="font-size:17px; color:rgba(255,255,255,0.6); max-width:680px; margin:0 auto; line-height:1.8;">
                自动化数据清洗 · 多维度分析 · 智能报告生成，为应急管理提供数字化引擎
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div style="display:flex; justify-content:center; gap:24px; flex-wrap:wrap; padding:10px 0 30px 0;">
            <div class="step-card">
                <span class="step-number">01</span>
                <div class="step-label">数据导入与清洗</div>
                <div class="step-desc">支持 Excel 批量上传，自动清洗</div>
            </div>
            <div class="step-card">
                <span class="step-number">02</span>
                <div class="step-label">多维度综合分析</div>
                <div class="step-desc">分灾种、分时段、分区域深度分析</div>
            </div>
            <div class="step-card">
                <span class="step-number">03</span>
                <div class="step-label">自动化可视化</div>
                <div class="step-desc">6 种图表类型，直观呈现数据</div>
            </div>
            <div class="step-card">
                <span class="step-number">04</span>
                <div class="step-label">智能报告生成</div>
                <div class="step-desc">一键生成 Word 分析报告</div>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
        <div style="text-align:center; padding:20px 0 10px 0; background:rgba(212,175,55,0.05); border-radius:16px; border:1px solid rgba(212,175,55,0.08);">
            <span style="color:rgba(255,255,255,0.4); font-size:14px;">
                ⚡ 企业命题：四川省减灾中心 &nbsp;·&nbsp; 数智应急 · 赋能决策
            </span>
        </div>
    """, unsafe_allow_html=True)

# ---------- 数据导入 ----------
elif st.session_state.page == '数据导入':
    st.markdown("## 📥 数据导入与清洗")
    st.markdown("上传符合模板格式的 Excel 灾情数据，系统将自动完成清洗与标准化。")
    
    uploaded_file = st.file_uploader("选择 Excel 文件 (.xlsx / .xls)", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            df_raw = pd.read_excel(uploaded_file, skiprows=1)
            with st.spinner("正在清洗数据..."):
                df_clean = clean_data(df_raw)
                save_data(df_clean)
            st.success(f"✅ 上传成功，共 {len(df_clean)} 条记录。")
            st.markdown("### 📋 数据预览 (前5行)")
            st.dataframe(df_clean.head(5), use_container_width=True)
        except Exception as e:
            st.error(f"❌ 读取文件失败: {str(e)}")
    else:
        st.info("📤 请上传 Excel 文件开始。")

# ---------- 多维度分析 ----------
elif st.session_state.page == '多维度分析':
    st.markdown("## 📊 综合分析仪表板")
    df = load_data()
    if df.empty:
        st.warning("⚠️ 暂无数据，请先上传灾情数据。")
    else:
        # 核心指标卡片
        stats = get_summary_stats(df)
        cols = st.columns(4)
        for i, (key, value) in enumerate(stats.items()):
            with cols[i % 4]:
                st.markdown(f"""
                    <div class="stat-card">
                        <div class="stat-label">{key}</div>
                        <div class="stat-value">{value}</div>
                    </div>
                """, unsafe_allow_html=True)
        
        st.markdown("---")
        
        # 图表布局
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📈 时间趋势（月）")
            trend = get_time_trend(df, 'M')
            if not trend.empty:
                fig = px.line(trend, x='时段', y='受灾人口(人)', title="受灾人口月度变化")
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white', showlegend=False)
                fig.update_traces(line=dict(color='#d4af37', width=2))
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无时间趋势数据。")
        
        with col2:
            st.markdown("#### 🥧 灾种损失占比")
            disaster = get_disaster_type_analysis(df)
            if not disaster.empty:
                fig = px.pie(disaster, values='直接经济损失(万元)', names='灾种', title="各灾种经济损失占比", color_discrete_sequence=['#d4af37', '#f7e68a', '#ff6b6b', '#4ecdc4', '#45b7d1'])
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无灾种数据。")
        
        col3, col4 = st.columns(2)
        with col3:
            st.markdown("#### 🗺️ 各区域经济损失")
            region = get_region_analysis(df)
            if not region.empty:
                fig = px.bar(region, x='区域', y='直接经济损失(万元)', title="各区域经济损失", color_discrete_sequence=['#d4af37'])
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无区域数据。")
        
        with col4:
            st.markdown("#### 🧩 损失结构")
            loss = get_loss_structure(df)
            if loss:
                loss_df = pd.DataFrame(list(loss.items()), columns=['类型', '占比(%)'])
                fig = px.pie(loss_df, values='占比(%)', names='类型', title="各类损失占比", color_discrete_sequence=['#d4af37', '#f7e68a', '#ff6b6b', '#4ecdc4', '#45b7d1', '#6c5ce7'])
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无损失结构数据。")
        
        # 自定义分析
        st.markdown("---")
        st.markdown("### 🔍 自定义深度分析")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🎯 高风险区域-灾种组合 (TOP10)")
            custom1 = get_custom_analysis1(df)
            if not custom1.empty:
                custom1['组合'] = custom1['区域'] + ' - ' + custom1['灾种']
                fig = px.bar(custom1, x='组合', y='直接经济损失(万元)', title="经济损失 TOP10 组合", color_discrete_sequence=['#d4af37'])
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无数据。")
        with c2:
            st.markdown("#### 🌡️ 月份-灾种发生频次热力图")
            custom2 = get_custom_analysis2(df)
            if not custom2.empty:
                melt_df = custom2.melt(id_vars='月份', var_name='灾种', value_name='频次')
                fig = px.density_heatmap(melt_df, x='月份', y='灾种', z='频次', title="各月份灾种频次", color_continuous_scale=['#1a2a4a', '#d4af37', '#f7e68a'])
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无数据。")

# ---------- 智能报告 ----------
elif st.session_state.page == '智能报告':
    st.markdown("## 📄 智能报告生成")
    df = load_data()
    if df.empty:
        st.warning("⚠️ 暂无数据，请先上传灾情数据。")
    else:
        st.markdown("点击下方按钮，系统将自动生成一份 **Word 格式** 的综合分析报告。")
        st.markdown("报告包含：总体概况、分灾种分析、损失结构、结论建议。")
        
        if st.button("🚀 生成并下载报告 (Word)", use_container_width=True):
            with st.spinner("⏳ 正在生成报告..."):
                file_stream = generate_report(df)
                st.download_button(
                    label="📥 点击下载报告",
                    data=file_stream,
                    file_name="灾情分析报告.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

# ==================== 页脚 ====================
st.markdown("""
    <div class="footer">
        <span>⚡ 企业命题：四川省减灾中心</span> &nbsp;|&nbsp; © 2026 灾智云 · 数智应急赋能平台
    </div>
""", unsafe_allow_html=True)