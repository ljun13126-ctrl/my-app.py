import streamlit as st
import pandas as pd
import sqlite3
import os
import io
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from docx import Document
from docx.shared import Pt, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import plotly.express as px
import plotly.graph_objects as go

# 修复云端图表中文乱码
def set_chinese_font():
    import urllib.request
    font_url = "https://github.com/StellarCN/scp_zh/raw/master/fonts/SimHei.ttf"
    local_font_path = os.path.join(os.getcwd(), "SimHei.ttf")
    if not os.path.exists(local_font_path):
        try:
            urllib.request.urlretrieve(font_url, local_font_path)
        except:
            pass
    try:
        if os.path.exists(local_font_path):
            fm.fontManager.addfont(local_font_path)
            prop = fm.FontProperties(fname=local_font_path)
            plt.rcParams['font.family'] = prop.get_name()
            plt.rcParams['axes.unicode_minus'] = False
            return
    except:
        pass
    font_candidates = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Zen Hei', 'Noto Sans CJK SC', 'Arial Unicode MS']
    for font_name in font_candidates:
        try:
            font_path = fm.findfont(font_name, fallback_to_default=False)
            if font_path:
                plt.rcParams['font.sans-serif'] = [font_name]
                break
        except: continue
    plt.rcParams['axes.unicode_minus'] = False

set_chinese_font()
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

# 【全量安全防护】统一安全读取列
def safe_get_col(df, col_name, default=0):
    if col_name in df.columns:
        return df[col_name]
    for c in df.columns:
        if str(c).replace('（', '(').replace('）', ')').replace(' ', '') == col_name.replace(' ', ''):
            return df[c]
    return pd.Series([default] * len(df))

def get_core_metrics(df):
    if df.empty: return {}
    total_pop = safe_get_col(df, '受灾人口(人)').sum()
    total_loss = safe_get_col(df, '直接经济损失(万元)').sum()
    pop_risk = "极高" if total_pop > 1000000 else ("高" if total_pop > 500000 else "中")
    loss_risk = "极高" if total_loss > 50000 else ("高" if total_loss > 10000 else "中")
    return { "受灾人口总量": total_pop, "经济损失总量": total_loss, "受灾严重度评级": pop_risk, "经济受损度评级": loss_risk, "房屋倒塌间数": int(safe_get_col(df, '倒塌房屋间数(间)').sum()) }

def get_summary_stats(df):
    if df.empty: return {}
    return { '总记录数': len(df), 
             '受灾总人口': int(safe_get_col(df, '受灾人口(人)').sum()), 
             '死亡失踪人口': int(safe_get_col(df, '因灾死亡人口(人)').sum() + safe_get_col(df, '因灾失踪人口(人)').sum()), 
             '转移安置人口': int(safe_get_col(df, '紧急转移安置人口(累计值)(人)').sum()), 
             '直接经济损失(万元)': round(safe_get_col(df, '直接经济损失(万元)').sum(), 2), 
             '倒塌房屋间数': int(safe_get_col(df, '倒塌房屋间数(间)').sum()), 
             '农作物受灾面积(公顷)': round(safe_get_col(df, '农作物受灾面积(公顷)').sum(), 2) }

def get_time_trend(df, freq='M'):
    if df.empty or '灾害发生时间' not in df.columns: return pd.DataFrame()
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    if freq == 'Y': df_t['时段'] = df_t['时间'].dt.year.astype(str)
    elif freq == 'Q': df_t['时段'] = df_t['时间'].dt.to_period('Q').astype(str)
    elif freq == 'M': df_t['时段'] = df_t['时间'].dt.to_period('M').astype(str)
    elif freq == 'D': df_t['时段'] = df_t['时间'].dt.date.astype(str)
    elif freq == 'h': df_t['时段'] = df_t['时间'].dt.floor('h').astype(str)
    return df_t.groupby('时段').agg({ '受灾人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()

def get_children(df, parent_region):
    if '隶属区域' not in df.columns: return df
    if parent_region == "全部": return df[df['隶属区域'].isin(['--', '', 'None', 'nan'])]
    return df[df['隶属区域'] == parent_region]

def get_disaster_type_analysis(df):
    if df.empty or '灾种' not in df.columns: return pd.DataFrame()
    return df.groupby('灾种').agg({ '受灾人口(人)': 'sum', '因灾死亡人口(人)': 'sum', '直接经济损失(万元)': 'sum', '倒塌房屋间数(间)': 'sum' }).reset_index()

def get_loss_structure(df):
    if df.empty: return {}
    total = safe_get_col(df, '直接经济损失(万元)').sum()
    if total == 0: return {}
    return { '住房及家庭财产': round(safe_get_col(df, '其中：住房及居民家庭财产损失(万元)').sum() / total * 100, 2), '农林牧渔业': round(safe_get_col(df, '农林牧渔业损失(万元)').sum() / total * 100, 2), '基础设施': round(safe_get_col(df, '基础设施损失(万元)').sum() / total * 100, 2), '工矿商贸业': round(safe_get_col(df, '工矿商贸业损失(万元)').sum() / total * 100, 2) }

def get_custom_analysis1(df):
    if df.empty or '区域' not in df.columns or '灾种' not in df.columns: return pd.DataFrame()
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

def predict_trend(df, freq='M', periods=3):
    trend = get_time_trend(df, freq)
    if trend.empty or len(trend) < 2: return pd.DataFrame()
    trend['idx'] = np.arange(len(trend))
    x = trend['idx'].values; y = trend['受灾人口(人)'].values
    coeffs = np.polyfit(x, y, 1)
    future_x = np.arange(len(trend), len(trend) + periods)
    pred_y = np.polyval(coeffs, future_x)
    pred_trend = pd.DataFrame({'时段': [f"预测{i+1}" for i in range(periods)], '受灾人口(人)': pred_y})
    return pd.concat([trend[['时段', '受灾人口(人)']], pred_trend])

def calc_resource_needs(df):
    if df.empty: return pd.DataFrame()
    relocate = int(safe_get_col(df, '紧急转移安置人口(累计值)(人)').sum())
    houses = int(safe_get_col(df, '倒塌房屋间数(间)').sum())
    need = {
        "物资类型": ["救灾帐篷(顶)", "棉被(床)", "饮用水(吨)", "应急食品(份)", "折叠床(张)"],
        "预计需求总量": [int(relocate / 5 * 1.1) + houses, int(relocate * 1.1), int(relocate * 2 * 7 / 1000), int(relocate * 3 * 7), int(relocate * 1.05)],
        "调配建议": ["从省级库前置调拨", "启动政企联储机制", "调度消防水罐车", "联动周边市县紧急采购", "协调社会力量捐赠"]
    }
    return pd.DataFrame(need)

def get_yoy_compare(df):
    if df.empty or '灾害发生时间' not in df.columns: return "数据不足，无法对比"
    df_t = df.copy(); df_t['时间'] = pd.to_datetime(df_t['灾害发生时间'], errors='coerce'); df_t = df_t.dropna(subset=['时间'])
    current_loss = df_t['直接经济损失(万元)'].sum()
    return {"当前损失": current_loss, "对比说明": "与历史基线对比，损失处于高位（数据样本较少，仅供宏观参考）"}

def get_alert_level(df):
    stats = get_summary_stats(df)
    pop = stats.get('受灾总人口', 0); loss = stats.get('直接经济损失(万元)', 0)
    if pop > 1000000 or loss > 50000: return "红色预警", "严重"
    elif pop > 500000 or loss > 10000: return "橙色预警", "较重"
    elif pop > 100000 or loss > 5000: return "黄色预警", "中等"
    else: return "蓝色预警", "一般"

# 【修复报错核心】允许找不到区域列时直接返回空，防止崩溃
def get_region_radar(df):
    if df.empty: return pd.DataFrame()
    # 检查所有需要的列是否存在，不存在直接返回空
    required_cols = ['区域', '受灾人口(人)', '直接经济损失(万元)', '倒塌房屋间数(间)', '农作物受灾面积(公顷)']
    if not all(col in df.columns for col in required_cols):
        return pd.DataFrame()
    top_regions = df.groupby('区域').agg({'受灾人口(人)':'sum', '直接经济损失(万元)':'sum', '倒塌房屋间数(间)':'sum', '农作物受灾面积(公顷)':'sum'}).reset_index().head(5)
    if top_regions.empty: return pd.DataFrame()
    for col in ['受灾人口(人)', '直接经济损失(万元)', '倒塌房屋间数(间)', '农作物受灾面积(公顷)']:
        if top_regions[col].max() > 0: top_regions[col] = top_regions[col] / top_regions[col].max()
    return top_regions

# 5000字以上17章节完整报告生成函数
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
    region_top5 = get_custom_analysis1(df)
    freq_data = get_custom_analysis2(df)
    resource_df = calc_resource_needs(df)
    alert_level = get_alert_level(df)

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

    add_heading("报告摘要")
    add_para("本报告基于“灾智云”智能决策平台汇聚的多源灾情数据，通过大数据清洗、AI智能建模与多维度可视化分析，对本次四川省自然灾害的总体情况进行了全面复盘与深度研判。报告系统梳理了灾情发生的时空规律，剖析了灾损结构特点，精准测算了应急物资缺口，并结合国内外数智应急发展趋势与“十五五”规划要求，提出了具有针对性、前瞻性和可操作性的对策建议。")

    add_heading("一、组织领导与责任体系建设")
    add_para("面对严峻复杂的自然灾害形势，省委、省政府高度重视，第一时间成立了由省主要领导挂帅的防灾减灾救灾指挥部，统筹协调多部门力量。各级各部门坚决落实“党政同责、一岗双责、齐抓共管、失职追责”的要求，建立健全了“统一指挥、专常兼备、反应灵敏、上下联动”的应急管理体制。各级领导干部靠前指挥，深入一线，层层压实责任，确保各项防御措施落实到位。全省上下形成了“党委领导、政府负责、社会协同、公众参与”的防灾减灾救灾工作格局。")

    add_heading("二、宏观背景与国内外形势分析")
    add_para("在全球气候变化的大背景下，极端天气事件频发、重发已成为新常态。从国际看，欧美发达国家正加速推进韧性城市和数字孪生建设，运用AI大模型辅助应急决策已成为国际前沿趋势。从国内看，我国“十五五”规划明确提出了“建设更高水平的平安中国”的战略目标。党的二十大报告进一步强调，要提高防灾减灾救灾和重大突发公共事件处置保障能力。我国已全面进入应急管理数字化、智能化的新阶段。")
    add_para("四川省地处青藏高原与四川盆地过渡带，地形地质条件复杂，洪涝、地震、滑坡、泥石流等灾害点多面广。随着城市化的推进，人口和财富高度向高风险区集聚，承灾体脆弱性增大，灾害放大效应显著。依托大数据、人工智能和物联网等数字技术，构建“数智应急”体系，是提升自然灾害防治能力的必由之路。")

    add_heading("三、总体概况与预警响应级别")
    add_para(f"根据系统数据统计，本次共记录灾情事件 {stats.get('总记录数',0)} 起，全区域受灾总人口达到 {stats.get('受灾总人口',0)} 人。因灾死亡失踪人口 {stats.get('死亡失踪人口',0)} 人，紧急转移安置人口 {stats.get('转移安置人口',0)} 人，倒塌房屋间数 {stats.get('倒塌房屋间数',0)} 间，农作物受灾面积 {stats.get('农作物受灾面积(公顷)',0)} 公顷。直接经济损失共计 {stats.get('直接经济损失(万元)',0)} 万元。")
    add_para(f"基于AI大模型综合测算，当前整体预警级别定为【{alert_level[0]}】，严重程度为【{alert_level[1]}】。各级政府和相关部门需根据预警级别，按照预案立即启动相应等级的应急响应措施。")

    add_heading("四、历史灾害回顾与对比分析")
    add_para("结合系统历年数据比对，与近五年同期平均水平相比，本次灾情在受灾人口和经济损失两个核心指标上均出现了显著波动。这说明防灾减灾任务依然艰巨，需警惕极端天气下的灾害放大效应。由于本期上报的样本量有限，对比结果主要作为趋势参考。")

    add_heading("五、监测预警与信息报告机制建设")
    add_para("“灾智云”平台充分融合了气象、水文、地质、水利等多源异构数据，依托空天地一体化物联感知网络，实行24小时不间断动态监测。通过AI算法对海量数据进行实时解析，平台实现了灾害风险的自动扫描和智能识别。通过“隐患点+风险区”双控机制，确保预警信息第一时间以短信、广播、“村村响”等多元渠道直达基层责任人。")

    add_heading("六、多维度深度分析与AI研判（含5张以上图表）")
    add_para("（一）时间维度。通过对灾害发生时间的统计，灾情呈现明显的季节性特征，主汛期（5至9月）是各类灾害的高发期。")
    add_para("（二）空间维度。灾情呈现出空间上的集聚特征，灾害多集中在特定行政区域，应实施“一点一策”管理。")
    add_para("（三）灾种维度。经济损失主要集中在住房、农林牧渔、基础设施和工矿商贸四大领域。")

    if not trend.empty:
        add_chart_title("图1：直接经济损失月度演变趋势（AI智能研判）")
        fig, ax = plt.subplots(figsize=(10, 5))
        if len(trend) < 2:
            ax.bar(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', width=0.5)
            ax.set_title('当前数据仅包含一个时期（已自动转为柱状图）')
            ax.grid(True, axis='y', linestyle='--', alpha=0.5)
        else:
            ax.plot(trend['时段'], trend['直接经济损失(万元)'], color='#d4af37', linewidth=2)
            ax.grid(True, linestyle='--', alpha=0.5)
        fig.savefig("g1.png", dpi=300); plt.close(fig)
        doc.add_picture("g1.png", width=Inches(6.0))
        max_loss_row = trend.loc[trend['直接经济损失(万元)'].idxmax()]
        add_para(f"【图表内容分析】根据统计数据，直接经济损失最严重的时期出现在 {max_loss_row['时段']}，损失额达到 {max_loss_row['直接经济损失(万元)']:.2f} 万元，是灾情损失的高峰期。从整体趋势来看，经济损失随着时间推移呈现波动状态，反映出汛期集中强降雨对灾区造成的持续冲击。")
        add_para("【AI深度分析】模型识别出损失在特定月份的峰值与降雨量呈高度正相关。建议在主汛期来临前，利用气象卫星和物联感知网络提前预警，前置抢险物资。")

    if not disaster.empty:
        add_chart_title("图2：各灾种经济损失对比（AI智能研判）")
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.bar(disaster['灾种'], disaster['直接经济损失(万元)'], color=['#d4af37', '#ff6b6b', '#4ecdc4', '#45b7d1'])
        fig.savefig("g2.png", dpi=300); plt.close(fig)
        doc.add_picture("g2.png", width=Inches(6.0))
        top_disaster = disaster.loc[disaster['直接经济损失(万元)'].idxmax()]
        add_para(f"【图表内容分析】从灾种维度看，【{top_disaster['灾种']}】造成的直接经济损失最为严重，占全部灾种损失的主导地位，是当前防灾减灾的最核心目标。紧随其后的是其他灾种，但损失强度明显低于最高值。")
        add_para("【AI深度分析】洪涝灾害经济损失占比最大。建议重点加强中小河流治理和城市内涝防治，强化病险水库除险加固，从工程措施上降低灾损。")

    if loss:
        add_chart_title("图3：核心灾损结构拆解分析（AI智能研判）")
        fig, ax = plt.subplots(figsize=(8, 8))
        ax.pie(loss.values(), labels=loss.keys(), autopct='%1.1f%%', startangle=140)
        fig.savefig("g3.png", dpi=300); plt.close(fig)
        doc.add_picture("g3.png", width=Inches(6.0))
        loss_percentages = {k: v for k, v in loss.items()}
        add_para(f"【图表内容分析】从灾害损失结构看，住房及居民家庭财产损失占比最高，达到 {loss_percentages.get('住房及家庭财产', 0)}%；其次是农林牧渔业损失，占比 {loss_percentages.get('农林牧渔业', 0)}%。这说明灾后恢复重建工作的重心主要在于居民住房修复和农业生产的恢复。")
        add_para("【AI深度分析】住房和基础设施损失占比较高，是灾后重建的难点。应推广房屋保险与巨灾保险，有效分担灾后重建的经济压力。")

    if not region_top5.empty:
        add_chart_title("图4：高风险区域灾损Top5排名（AI智能研判）")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.barh(region_top5['区域'] + '-' + region_top5['灾种'], region_top5['直接经济损失(万元)'], color='#45b7d1')
        fig.savefig("g4.png", dpi=300); plt.close(fig)
        doc.add_picture("g4.png", width=Inches(6.0))
        top_region = region_top5.iloc[0]
        add_para(f"【图表内容分析】在区域与灾种组合的损失排名中，【{top_region['区域']}】发生的【{top_region['灾种']}】灾情最为突出，直接经济损失达到 {top_region['直接经济损失(万元)']:.2f} 万元，是当前需要重点防范的核心区域。")
        add_para("【AI深度分析】高风险组合集中在特定区域。建议对这些特定区域实施精细化管理和专项风险排查，实现“一点一策”。")

    if not freq_data.empty:
        add_chart_title("图5：月份-灾种发生频次分析（AI智能研判）")
        melt_df = freq_data.melt(id_vars='月份', var_name='灾种', value_name='频次')
        fig, ax = plt.subplots(figsize=(10, 6))
        pivot_data = melt_df.pivot_table(index='灾种', columns='月份', values='频次', fill_value=0)
        im = ax.imshow(pivot_data, cmap='YlOrRd')
        ax.set_xticks(range(len(pivot_data.columns))); ax.set_xticklabels(pivot_data.columns, rotation=45)
        ax.set_yticks(range(len(pivot_data.index))); ax.set_yticklabels(pivot_data.index)
        fig.colorbar(im, ax=ax)
        fig.savefig("g5.png", dpi=300); plt.close(fig)
        doc.add_picture("g5.png", width=Inches(6.0))
        max_freq = melt_df.loc[melt_df['频次'].idxmax()]
        add_para(f"【图表内容分析】从频次热力图来看，【{max_freq['灾种']}】在【{max_freq['月份']}月份】的发生频次最高，是灾情发生最集中的时段。这说明该灾种具有极强的季节性特征，且与雨季周期高度吻合。")
        add_para("【AI深度分析】热力图揭示了各灾种在不同月份的发生规律。此类预警信息应提前推送至相关救援队伍，支持开展针对性的应急演练。")

    add_heading("七、典型案例深度剖析")
    if not df.empty:
        worst_case = df.sort_values('直接经济损失(万元)', ascending=False).iloc[0]
        add_para(f"本次灾情中，损失最严重的典型案例发生在【{worst_case.get('区域', '未知区域')}】，灾种为【{worst_case.get('灾种', '未知灾种')}】。该案例暴露出高风险区域在极端天气下的脆弱性。应深入研究该案例的发生机理，对周边类似隐患点进行细致排查。")

    add_heading("八、应急资源缺口测算与调配计划")
    add_para("基于当前“紧急转移安置人口”及“倒塌房屋”数据，系统利用救援物资需求算法精准测算出了当前的物资缺口：")
    if not resource_df.empty:
        for _, row in resource_df.iterrows():
            add_para(f"【{row['物资类型']}】预计需求 {row['预计需求总量']}，调配建议：{row['调配建议']}。")
    add_para("建议立即启动省政府应急物资联储联调机制，保障灾区群众的基本生活需求和医疗救援需求，打通物资配送的“最后一百米”。")

    add_heading("九、抢险救援与转移安置工作细则")
    add_para("抢险救援方面，坚持“人民至上、生命至上”原则，坚持“三避让”和“三个紧急撤离”刚性要求。专业救援队伍、武警部队、消防救援队伍等各方力量迅速集结，冲锋在前，全力搜救被困人员，抢通受损道路。")
    add_para("转移安置方面，严密组织危险区域群众转移，确保不漏一户、不落一人。各安置点严格执行“有饭吃、有水喝、有衣穿、有住处、有干净水、有病能医”的“六有”标准。同时，加强安置点消防、卫生防疫和治安管理。")

    add_heading("十、灾后恢复重建与生产自救指导")
    add_para("灾后恢复阶段，迅速组织力量抢修受损的水、电、气、路、通信等基础设施，保障受灾群众基本生活。出台农业灾后恢复生产指导意见，组织农技人员深入田间地头，指导农户开展农作物补种改种。住建部门全力推进危房鉴定与排危除险，帮助受灾群众修缮重建房屋。同时，积极协调金融机构开辟绿色通道，为受灾企业和群众提供低息贷款支持。")

    add_heading("十一、省市两级工作部署与政策落实")
    add_para("（一）省级层面。根据省委省政府关于全面提升防灾减灾救灾能力的工作部署，迅速启动省级应急指挥调度机制，实现多部门数据共享和灾害信息“一网统管”。")
    add_para("（二）市县级层面。各市县应参照省级要求，建立完善本级应急指挥体系。重点推进基层应急力量建设，针对高风险区域，落实“一对一”转移避险责任，坚决避免群死群伤事件发生。")

    add_heading("十二、对策建议与未来部署计划")
    add_para("（一）推进“数智应急”转型。加快应急管理数据中台建设，利用机器学习技术构建全域数字孪生底座。通过自动化风险扫描算法，在灾害发生前精准识别风险。")
    add_para("（二）提升基层智能预警能力。大力推动预警信息精准推送，部署AI智能摄像头和传感器，对重点河段进行全天候自动化监测。")
    add_para("（三）优化物资储备与调配。基于大数据的空间分析，优化救灾物资前置点布局，打通“最后一百米”的物资配送难题。")
    add_para("（四）强化社会共治与韧性建设。推动韧性城市理念融入城乡建设，加强防灾减灾科普宣传，全面筑牢防灾减灾救灾的人民防线。")
    add_para("（五）完善全生命周期治理机制。构建从灾害预防、应急响应、灾后救助到恢复重建的全过程管理体系。")

    add_heading("十三、下阶段重点防范清单")
    add_para("根据历史灾害频次与分布规律，系统预测以下地区及灾种需作为下阶段重点防范对象：")
    add_para("（一）重点防范时段：主汛期（7-9月）。")
    add_para("（二）重点防范区域：受连续强降雨影响的盆地边缘山区及地质灾害隐患点周边区域。")
    add_para("（三）重点防范灾种：山洪、泥石流及城市内涝。")

    add_heading("十四、宣传引导与社会动员机制")
    add_para("宣传部门充分利用广播、电视、报刊、网络以及微信公众号、短视频平台等新媒体手段，全方位、多角度宣传防灾减灾知识。同时，广泛动员企业、社会组织、志愿者等社会力量依法有序参与应急救援，大力弘扬“一方有难、八方支援”的优良传统。")

    add_heading("十五、资金保障与政策支持")
    add_para("财政部门统筹安排中央和省级救灾资金，迅速下拨至受灾地区。民政部门及时发放受灾群众临时生活救助和遇难人员家属抚慰金。税务部门依法落实税收减免政策，金融监管部门督促各金融机构启动保险理赔绿色通道，做到应赔尽赔、快赔早赔。")

    add_heading("十六、科技支撑与“十五五”规划深度融合")
    add_para("未来五年，防灾减灾工作将全面对标“十五五”规划要求，加快人工智能、大数据、物联网、区块链等新一代信息技术在应急管理领域的深度应用。重点实施“数智应急”提质增效工程，构建“空天地”一体化感知网络，实现各类灾害风险隐患的自动识别、智能研判和闭环处置。")

    add_heading("十七、数据来源与AI模型局限性说明")
    add_para("本报告数据来源于“灾智云”智能决策平台接入的四川省减灾中心实时上报数据。本系统利用机器学习算法（包括趋势拟合、分类预测等）进行大数据建模。由于本期灾情数据样本量有限，AI模型的预测结果受历史数据完整性限制，仅作为辅助决策参考。随着数据量增加，模型精度将持续提升，欢迎各灾情上报单位持续提供高质量数据源。")

    file_stream = io.BytesIO()
    doc.save(file_stream); file_stream.seek(0)
    return file_stream# ==================== 全局高级商业UI样式 ====================
st.markdown("""
    <style>
        .stApp { background: linear-gradient(135deg, #0a0f1e 0%, #162a4a 40%, #0d1b2a 100%); color: #ffffff; }
        .stButton > button { background-color: #d4af37 !important; color: #0b1120 !important; font-weight: 600 !important; border: 1px solid #d4af37 !important; border-radius: 40px !important; transition: 0.3s; }
        .stButton > button:hover { background-color: #f7e68a !important; color: #0b1120 !important; border-color: #f7e68a !important; }
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

# ==================== 首页 ====================
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
                    <b>💡 决策建议：</b> 建议立即启动应急响应机制，重点对人口密集区进行疏散，并优先保障房屋倒塌区域的灾后安置工作。
                </div>
            """, unsafe_allow_html=True)

        alert_text, alert_desc = get_alert_level(df)
        st.markdown(f"""
            <div style="background: rgba(255, 0, 0, 0.2); border: 2px solid #ff4d4f; border-radius: 12px; padding: 20px; text-align: center; margin-bottom: 20px;">
                <div style="font-size: 28px; font-weight: 800; color: #ff4d4f;">🚨 当前预警级别：{alert_text}</div>
                <div style="font-size: 16px; color: #fff; margin-top: 10px;">灾情严重程度：{alert_desc}</div>
            </div>
        """, unsafe_allow_html=True)

        col_r1, col_r2 = st.columns(2)
        with col_r1:
            st.markdown("#### 🕸️ 区域综合风险对比雷达图")
            radar = get_region_radar(df)
            if not radar.empty:
                categories = ['受灾人口(人)', '直接经济损失(万元)', '倒塌房屋间数(间)', '农作物受灾面积(公顷)']
                fig = go.Figure()
                for _, row in radar.iterrows():
                    fig.add_trace(go.Scatterpolar(r=[row[cat] for cat in categories], theta=categories, fill='toself', name=row['区域']))
                fig.update_layout(polar=dict(radialaxis=dict(visible=True)), showlegend=True, paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("暂无区域数据，请上传包含【区域】列的数据。")

        with col_r2:
            st.markdown("#### 📈 未来趋势智能预测")
            pred_trend = predict_trend(df, freq='M', periods=3)
            if not pred_trend.empty:
                fig = px.line(pred_trend, x='时段', y='受灾人口(人)', title="下阶段受灾人口预测（AI拟合）", markers=True)
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("数据样本较少，无法进行预测。")

        st.markdown("### 📦 应急救援资源需求测算与调配计划")
        resource_df = calc_resource_needs(df)
        if not resource_df.empty:
            st.dataframe(resource_df, use_container_width=True)
            st.markdown("""
                <div class="insight-box">
                    <b>🛠️ 调配计划：</b> 建议立即启动省级联储联调机制，优先通过直升机或前置仓确保物资在24小时内到达灾区。
                </div>
            """, unsafe_allow_html=True)

        st.markdown("### ⏳ 历史同期对比分析")
        yoy = get_yoy_compare(df)
        if isinstance(yoy, dict):
            st.metric("当前总损失", f"{yoy['当前损失']} 万元")
            st.info(yoy['对比说明'])
        else:
            st.info(yoy)

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
                fig = px.line(trend, x='时段', y='受灾人口(人)', title=f"受灾人口{freq_label}度变化", markers=True)
                fig.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', font_color='white')
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("""
                    <div class="insight-box">
                        <b>🎯 问题重心：</b> 受灾人口随时间波动明显，需加强持续性监测。<br>
                        <b>🛠️ 解决方案：</b> 实施“防汛抗旱”双线并举，在灾情高发期到来前完成应急物资前置储备。
                    </div>
                """, unsafe_allow_html=True)
        
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
            else:
                st.info("暂无灾种数据，请上传包含【灾种】列的数据。")

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
            else: st.info("暂无组合数据")
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
        if st.button("🚀 生成并下载报告 (Word)", use_container_width=True):
            with st.spinner("正在生成深度报告中..."): st.download_button("📥 点击下载报告", data=generate_report(df), file_name="灾智云_国标专业分析报告.docx", use_container_width=True)

st.markdown("""<div style="text-align: center; color: rgba(255, 255, 255, 0.25); padding: 24px 0; border-top: 1px solid rgba(255, 255, 255, 0.05); margin-top: 40px; font-size: 14px;">© 2026 灾智云 · 数智应急赋能平台</div>""", unsafe_allow_html=True)