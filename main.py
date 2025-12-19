# -*- coding: utf-8 -*-
"""
生产排产交付能力验证系统 - 主程序
"""
import sys
import os
from datetime import datetime
import io

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# 添加当前目录到路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import Config
from data_loader import OrderLoader, BOMLoader, CapacityLoader
from calculator import MaterialRequirementCalculator, ProductionScheduler
from analyzer import DeliveryAnalyzer, BottleneckDetector
from reporter import ReportGenerator


def print_banner():
    """打印系统横幅"""
    print("=" * 70)
    print(" " * 15 + "生产排产交付能力验证系统")
    print(" " * 20 + "Production Scheduling System")
    print("=" * 70)
    print()


def print_section(title):
    """打印章节标题"""
    print(f"\n{'=' * 70}")
    print(f"  {title}")
    print(f"{'=' * 70}")


def main():
    """主函数"""
    try:
        # 打印横幅
        print_banner()
        
        # 确保目录存在
        Config.ensure_directories()
        
        # ========== 1. 加载数据 ==========
        print_section("[1/6] 加载数据")
        
        # 加载订单数据
        print("  加载订单数据...")
        order_loader = OrderLoader(Config.get_input_file_path(Config.ORDERS_FILE))
        orders_df = order_loader.load()
        order_summary = order_loader.get_summary()
        print(f"  ✓ 订单数据加载完成")
        for key, value in order_summary.items():
            print(f"    - {key}: {value}")
        
        # 加载BOM数据
        print("\n  加载BOM数据...")
        bom_loader = BOMLoader(Config.get_input_file_path(Config.BOM_FILE))
        bom_df = bom_loader.load()
        bom_summary = bom_loader.get_summary()
        print(f"  ✓ BOM数据加载完成")
        for key, value in bom_summary.items():
            print(f"    - {key}: {value}")
        
        # 加载产能数据
        print("\n  加载产能数据...")
        capacity_loader = CapacityLoader(Config.get_input_file_path(Config.CAPACITY_FILE))
        capacity_df = capacity_loader.load()
        capacity_summary = capacity_loader.get_summary()
        print(f"  ✓ 产能数据加载完成")
        for key, value in capacity_summary.items():
            print(f"    - {key}: {value}")
        
        # ========== 2. 计算物料需求 ==========
        print_section("[2/6] 计算物料需求计划(MRP)")
        
        mrp_calculator = MaterialRequirementCalculator(orders_df, bom_df)
        mrp_df = mrp_calculator.calculate()
        mrp_summary = mrp_calculator.get_summary()
        
        print(f"  ✓ 物料需求计算完成")
        for key, value in mrp_summary.items():
            print(f"    - {key}: {value}")
        
        # ========== 3. 执行排产 ==========
        print_section("[3/6] 执行产能排产")
        
        start_date = orders_df['生产开工日期'].min()
        print(f"  排产开工日期: {start_date.strftime('%Y-%m-%d')}")
        
        scheduler = ProductionScheduler(mrp_df, capacity_df, start_date)
        schedule_df = scheduler.schedule()
        schedule_summary = scheduler.get_summary()
        
        print(f"  ✓ 产能排产完成")
        for key, value in schedule_summary.items():
            print(f"    - {key}: {value}")
        
        # ========== 4. 分析交付能力 ==========
        print_section("[4/6] 分析交付能力")
        
        analyzer = DeliveryAnalyzer(orders_df, schedule_df, bom_df, mrp_df)
        delivery_analysis = analyzer.analyze()
        analysis_summary = analyzer.get_summary(delivery_analysis)
        
        print(f"  ✓ 交付能力分析完成")
        for key, value in analysis_summary.items():
            print(f"    - {key}: {value}")
        
        # ========== 5. 识别瓶颈 ==========
        print_section("[5/6] 识别产能瓶颈")
        
        bottleneck_detector = BottleneckDetector(schedule_df, capacity_df)
        capacity_gap = bottleneck_detector.calculate_gap()
        bottleneck_summary = bottleneck_detector.summarize()
        
        print(f"  ✓ 瓶颈识别完成")
        print(f"    - 产能缺口物料数: {len(capacity_gap)}")
        print(f"    - 瓶颈物料总数: {len(bottleneck_summary)}")
        
        # ========== 6. 生成报告 ==========
        print_section("[6/6] 生成分析报告")
        
        report_path = Config.get_output_file_path(Config.REPORT_FILE)
        
        # 汇总统计信息
        summary_stats = {
            '订单总数': analysis_summary['订单总数'],
            '按时交付订单数': analysis_summary['按时交付订单数'],
            '延期订单数': analysis_summary['延期订单数'],
            '按时交付率(%)': analysis_summary['按时交付率'],
            '红色预警数': analysis_summary['红色预警数'],
            '黄色预警数': analysis_summary['黄色预警数'],
            '物料总数': mrp_summary['物料总数'],
            '瓶颈物料数': len(bottleneck_summary),
            '平均产能利用率(%)': schedule_summary['平均产能利用率']
        }
        
        reporter = ReportGenerator(report_path)
        reporter.generate(
            delivery_analysis=delivery_analysis,
            capacity_gap=capacity_gap,
            bottleneck_summary=bottleneck_summary,
            summary_stats=summary_stats
        )
        
        # ========== 输出关键预警 ==========
        print_section("关键预警信息")
        
        alerts = analyzer.get_alerts(delivery_analysis)
        
        # 红色预警
        if alerts['red']:
            print(f"\n🔴 红色预警 ({len(alerts['red'])}个订单延期>=7天):")
            for alert in alerts['red'][:5]:  # 只显示前5个
                print(f"  - {alert['订单号']}: 延期{alert['延期天数']}天, 瓶颈: {alert['瓶颈物料']}")
            if len(alerts['red']) > 5:
                print(f"  ... 还有{len(alerts['red']) - 5}个红色预警订单")
        else:
            print("\n🟢 无红色预警订单")
        
        # 黄色预警
        if alerts['yellow']:
            print(f"\n🟡 黄色预警 ({len(alerts['yellow'])}个订单延期1-6天):")
            for alert in alerts['yellow'][:5]:  # 只显示前5个
                print(f"  - {alert['订单号']}: 延期{alert['延期天数']}天, 瓶颈: {alert['瓶颈物料']}")
            if len(alerts['yellow']) > 5:
                print(f"  ... 还有{len(alerts['yellow']) - 5}个黄色预警订单")
        else:
            print("\n🟢 无黄色预警订单")
        
        # TOP瓶颈物料
        if not bottleneck_summary.empty:
            print(f"\n📊 TOP 5 瓶颈物料:")
            top_bottlenecks = bottleneck_detector.get_top_bottlenecks(bottleneck_summary, 5)
            for idx, (_, bottleneck) in enumerate(top_bottlenecks.iterrows(), 1):
                print(f"  {idx}. {bottleneck['物料编码']} - {bottleneck['瓶颈类型']} "
                      f"(延期{bottleneck['延期天数']}天, 利用率{bottleneck['产能利用率(%)']}%)")
        
        # ========== 完成 ==========
        print_section("分析完成")
        
        print(f"\n✓ 报告已生成: {report_path}")
        print(f"✓ 分析时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("\n" + "=" * 70)
        print(" " * 20 + "感谢使用本系统！")
        print("=" * 70)
        
        return 0
        
    except FileNotFoundError as e:
        print(f"\n❌ 文件错误: {str(e)}")
        print(f"\n请确保以下文件存在于 {Config.INPUT_DIR} 目录:")
        print(f"  - {Config.ORDERS_FILE}")
        print(f"  - {Config.BOM_FILE}")
        print(f"  - {Config.CAPACITY_FILE}")
        return 1
        
    except ValueError as e:
        print(f"\n❌ 数据验证错误: {str(e)}")
        print("\n请检查输入数据的格式和内容是否正确。")
        return 1
        
    except Exception as e:
        print(f"\n❌ 系统错误: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
