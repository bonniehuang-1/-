# -*- coding: utf-8 -*-
"""
生产排产交付能力验证系统 - Streamlit Web界面
"""
import streamlit as st
import pandas as pd
import sys
import os
from datetime import datetime
import io

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import Config
from data_loader import OrderLoader, BOMLoader, CapacityLoader
from calculator import MaterialRequirementCalculator, ProductionScheduler
from analyzer import DeliveryAnalyzer, BottleneckDetector
from reporter import ReportGenerator


# 页面配置
st.set_page_config(
    page_title="生产排产系统",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 自定义CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        padding: 1rem 0;
    }
    .metric-card {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin: 0.5rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 1rem;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 1rem;
        margin: 1rem 0;
    }
    .danger-box {
        background-color: #f8d7da;
        border-left: 4px solid #dc3545;
        padding: 1rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


def main():
    """主函数"""
    
    # 标题
    st.markdown('<div class="main-header">📊 生产排产交付能力验证系统</div>', unsafe_allow_html=True)
    st.markdown("---")
    
    # 侧边栏
    with st.sidebar:
        st.header("📁 数据上传")
        st.markdown("请上传以下3个Excel文件：")
        
        orders_file = st.file_uploader("1️⃣ 订单数据 (orders.xlsx)", type=['xlsx', 'xls'])
        bom_file = st.file_uploader("2️⃣ BOM数据 (bom.xlsx)", type=['xlsx', 'xls'])
        capacity_file = st.file_uploader("3️⃣ 产能数据 (capacity.xlsx)", type=['xlsx', 'xls'])
        
        st.markdown("---")
        
        # 使用示例数据选项
        use_sample = st.checkbox("使用示例数据", value=False)
        
        st.markdown("---")
        
        # 分析按钮
        analyze_button = st.button("🚀 开始分析", type="primary", use_container_width=True)
    
    # 主内容区
    if analyze_button:
        if use_sample or (orders_file and bom_file and capacity_file):
            run_analysis(orders_file, bom_file, capacity_file, use_sample)
        else:
            st.error("❌ 请上传所有必需的文件或选择使用示例数据")
    else:
        show_welcome()


def show_welcome():
    """显示欢迎页面"""
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("### 📤 上传数据\n在左侧上传订单、BOM和产能数据文件")
    
    with col2:
        st.info("### 🔍 分析处理\n系统自动计算物料需求和排产计划")
    
    with col3:
        st.info("### 📊 查看结果\n获取详细的分析报告和预警信息")
    
    st.markdown("---")
    
    # 功能介绍
    st.header("✨ 系统功能")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        **核心功能：**
        - ✅ 多层级BOM自动展开
        - ✅ 物料需求计划（MRP）计算
        - ✅ 产能排产模拟
        - ✅ 交付能力分析
        """)
    
    with col2:
        st.markdown("""
        **分析输出：**
        - 📈 交付状态统计
        - 🎯 瓶颈物料识别
        - ⚠️ 延期预警信息
        - 📥 Excel报告下载
        """)
    
    st.markdown("---")
    
    # 数据格式说明
    with st.expander("📋 数据格式说明"):
        st.markdown("""
        **订单数据必填列：**
        - 订单号、产品型号、数量、生产开工日期、发货日期
        
        **BOM数据必填列：**
        - 父物料编码、子物料编码、用量、层级、生产周期(天)
        
        **产能数据必填列：**
        - 物料编码、日产能上限
        """)


def run_analysis(orders_file, bom_file, capacity_file, use_sample):
    """运行分析"""
    
    try:
        # 确保目录存在
        Config.ensure_directories()
        
        # 进度条
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # 1. 加载数据
        status_text.text("📂 正在加载数据...")
        progress_bar.progress(10)
        
        if use_sample:
            # 使用示例数据
            orders_df = pd.read_excel(Config.get_input_file_path(Config.ORDERS_FILE))
            bom_df = pd.read_excel(Config.get_input_file_path(Config.BOM_FILE))
            capacity_df = pd.read_excel(Config.get_input_file_path(Config.CAPACITY_FILE))
        else:
            # 使用上传的文件
            orders_df = pd.read_excel(orders_file)
            bom_df = pd.read_excel(bom_file)
            capacity_df = pd.read_excel(capacity_file)
        
        # 数据验证和转换
        orders_df = validate_orders(orders_df)
        bom_df = validate_bom(bom_df)
        capacity_df = validate_capacity(capacity_df)
        
        progress_bar.progress(20)
        
        # 2. 计算MRP
        status_text.text("🔢 正在计算物料需求...")
        mrp_calculator = MaterialRequirementCalculator(orders_df, bom_df)
        mrp_df = mrp_calculator.calculate()
        progress_bar.progress(40)
        
        # 3. 执行排产
        status_text.text("📅 正在执行产能排产...")
        start_date = orders_df['生产开工日期'].min()
        scheduler = ProductionScheduler(mrp_df, capacity_df, start_date)
        schedule_df = scheduler.schedule()
        progress_bar.progress(60)
        
        # 4. 分析交付能力
        status_text.text("📊 正在分析交付能力...")
        analyzer = DeliveryAnalyzer(orders_df, schedule_df, bom_df, mrp_df)
        delivery_analysis = analyzer.analyze()
        progress_bar.progress(80)
        
        # 5. 识别瓶颈
        status_text.text("🎯 正在识别瓶颈...")
        bottleneck_detector = BottleneckDetector(schedule_df, capacity_df)
        capacity_gap = bottleneck_detector.calculate_gap()
        bottleneck_summary = bottleneck_detector.summarize()
        progress_bar.progress(90)
        
        # 6. 生成报告
        status_text.text("📝 正在生成报告...")
        report_path = Config.get_output_file_path(Config.REPORT_FILE)
        
        summary_stats = {
            '订单总数': len(delivery_analysis),
            '按时交付订单数': len(delivery_analysis[delivery_analysis['能否按时交付']]),
            '延期订单数': len(delivery_analysis[~delivery_analysis['能否按时交付']]),
            '物料总数': len(mrp_df),
            '瓶颈物料数': len(bottleneck_summary)
        }
        
        reporter = ReportGenerator(report_path)
        reporter.generate(delivery_analysis, capacity_gap, bottleneck_summary, summary_stats)
        
        progress_bar.progress(100)
        status_text.text("✅ 分析完成！")
        
        # 显示结果
        display_results(delivery_analysis, schedule_df, bottleneck_summary, 
                       analyzer, report_path)
        
    except Exception as e:
        st.error(f"❌ 分析过程中出现错误：{str(e)}")
        st.exception(e)


def validate_orders(df):
    """验证订单数据"""
    df['生产开工日期'] = pd.to_datetime(df['生产开工日期'])
    df['发货日期'] = pd.to_datetime(df['发货日期'])
    df['数量'] = df['数量'].astype(int)
    return df


def validate_bom(df):
    """验证BOM数据"""
    df['用量'] = df['用量'].astype(float)
    df['层级'] = df['层级'].astype(int)
    df['生产周期(天)'] = df['生产周期(天)'].astype(int)
    return df


def validate_capacity(df):
    """验证产能数据"""
    df['日产能上限'] = df['日产能上限'].astype(int)
    return df


def display_results(delivery_df, schedule_df, bottleneck_df, analyzer, report_path):
    """显示分析结果"""
    
    st.success("✅ 分析完成！")
    
    # 汇总统计
    st.header("📊 分析结果汇总")
    
    summary = analyzer.get_summary(delivery_df)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("订单总数", summary['订单总数'])
    
    with col2:
        st.metric("按时交付", summary['按时交付订单数'], 
                 delta=f"{summary['按时交付率']:.1f}%")
    
    with col3:
        st.metric("延期订单", summary['延期订单数'],
                 delta=f"-{summary['延期订单数']}" if summary['延期订单数'] > 0 else "0",
                 delta_color="inverse")
    
    with col4:
        st.metric("瓶颈物料", len(bottleneck_df))
    
    st.markdown("---")
    
    # 预警信息
    st.header("⚠️ 预警信息")
    
    alerts = analyzer.get_alerts(delivery_df)
    
    col1, col2 = st.columns(2)
    
    with col1:
        if alerts['red']:
            st.markdown('<div class="danger-box">', unsafe_allow_html=True)
            st.markdown(f"### 🔴 红色预警 ({len(alerts['red'])}个)")
            for alert in alerts['red'][:5]:
                st.write(f"- {alert['订单号']}: 延期{alert['延期天数']}天")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.markdown("### 🟢 无红色预警")
            st.markdown('</div>', unsafe_allow_html=True)
    
    with col2:
        if alerts['yellow']:
            st.markdown('<div class="warning-box">', unsafe_allow_html=True)
            st.markdown(f"### 🟡 黄色预警 ({len(alerts['yellow'])}个)")
            for alert in alerts['yellow'][:5]:
                st.write(f"- {alert['订单号']}: 延期{alert['延期天数']}天")
            st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.markdown("### 🟢 无黄色预警")
            st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown("---")
    
    # 详细数据表格
    tab1, tab2, tab3 = st.tabs(["📋 订单交付状态", "🎯 瓶颈物料", "📈 排产计划"])
    
    with tab1:
        st.subheader("订单交付状态")
        display_df = delivery_df[['订单号', '产品型号', '数量', '要求交付日期', 
                                   '预计完成日期', '延期天数', '状态', '瓶颈物料']].copy()
        
        # 格式化日期
        display_df['要求交付日期'] = display_df['要求交付日期'].dt.strftime('%Y-%m-%d')
        display_df['预计完成日期'] = display_df['预计完成日期'].dt.strftime('%Y-%m-%d')
        
        st.dataframe(display_df, use_container_width=True)
    
    with tab2:
        st.subheader("瓶颈物料分析")
        if not bottleneck_df.empty:
            display_bottleneck = bottleneck_df[['物料编码', '瓶颈类型', '日产能', 
                                                '总需求量', '产能利用率(%)', '延期天数']].copy()
            st.dataframe(display_bottleneck, use_container_width=True)
        else:
            st.info("未发现瓶颈物料")
    
    with tab3:
        st.subheader("排产计划")
        display_schedule = schedule_df[['物料编码', '总需求量', '日产能', 
                                        '开工日期', '预计完成日期', '延期天数', 
                                        '平均产能利用率']].copy()
        display_schedule['平均产能利用率'] = (display_schedule['平均产能利用率'] * 100).round(2)
        st.dataframe(display_schedule, use_container_width=True)
    
    st.markdown("---")
    
    # 下载报告
    st.header("📥 下载报告")
    
    with open(report_path, 'rb') as f:
        st.download_button(
            label="📊 下载完整Excel报告",
            data=f,
            file_name="delivery_analysis_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )


if __name__ == "__main__":
    main()
