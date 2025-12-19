# -*- coding: utf-8 -*-
"""
示例数据生成脚本
运行此脚本可以在input目录生成示例数据文件，用于测试系统
"""
import pandas as pd
import os
import sys
from datetime import datetime, timedelta

# 设置输出编码为UTF-8
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def create_sample_data():
    """创建示例数据文件"""
    
    # 确保input目录存在
    input_dir = "input"
    if not os.path.exists(input_dir):
        os.makedirs(input_dir)
        print(f"创建目录: {input_dir}")
    
    # ========== 1. 创建订单数据 ==========
    print("\n创建订单数据...")
    
    base_date = datetime(2024, 1, 15)
    
    orders_data = [
        {
            '订单号': 'ORD-2024-001',
            '产品型号': 'PROD-A100',
            '数量': 1000,
            '生产开工日期': base_date,
            '发货日期': base_date + timedelta(days=30),
            '优先级': 1
        },
        {
            '订单号': 'ORD-2024-002',
            '产品型号': 'PROD-A100',
            '数量': 500,
            '生产开工日期': base_date + timedelta(days=5),
            '发货日期': base_date + timedelta(days=35),
            '优先级': 2
        },
        {
            '订单号': 'ORD-2024-003',
            '产品型号': 'PROD-B200',
            '数量': 800,
            '生产开工日期': base_date + timedelta(days=3),
            '发货日期': base_date + timedelta(days=25),
            '优先级': 1
        },
        {
            '订单号': 'ORD-2024-004',
            '产品型号': 'PROD-A100',
            '数量': 600,
            '生产开工日期': base_date + timedelta(days=10),
            '发货日期': base_date + timedelta(days=40),
            '优先级': 3
        },
        {
            '订单号': 'ORD-2024-005',
            '产品型号': 'PROD-B200',
            '数量': 400,
            '生产开工日期': base_date + timedelta(days=7),
            '发货日期': base_date + timedelta(days=28),
            '优先级': 2
        }
    ]
    
    orders_df = pd.DataFrame(orders_data)
    orders_file = os.path.join(input_dir, 'orders.xlsx')
    orders_df.to_excel(orders_file, index=False)
    print(f"[OK] 订单数据已创建: {orders_file}")
    print(f"  - 订单数量: {len(orders_df)}")
    
    # ========== 2. 创建BOM数据 ==========
    print("\n创建BOM数据...")
    
    bom_data = [
        # PROD-A100的BOM
        {
            '父物料编码': 'PROD-A100',
            '父物料名称': 'A100成品',
            '子物料编码': 'SEMI-X100',
            '子物料名称': 'X100半成品',
            '用量': 2.0,
            '层级': 1,
            '生产周期(天)': 5
        },
        {
            '父物料编码': 'PROD-A100',
            '父物料名称': 'A100成品',
            '子物料编码': 'SEMI-Y100',
            '子物料名称': 'Y100半成品',
            '用量': 1.5,
            '层级': 1,
            '生产周期(天)': 3
        },
        # SEMI-X100的BOM
        {
            '父物料编码': 'SEMI-X100',
            '父物料名称': 'X100半成品',
            '子物料编码': 'RAW-M001',
            '子物料名称': 'M001原材料',
            '用量': 3.0,
            '层级': 2,
            '生产周期(天)': 2
        },
        {
            '父物料编码': 'SEMI-X100',
            '父物料名称': 'X100半成品',
            '子物料编码': 'RAW-M002',
            '子物料名称': 'M002原材料',
            '用量': 1.0,
            '层级': 2,
            '生产周期(天)': 2
        },
        # SEMI-Y100的BOM
        {
            '父物料编码': 'SEMI-Y100',
            '父物料名称': 'Y100半成品',
            '子物料编码': 'RAW-M002',
            '子物料名称': 'M002原材料',
            '用量': 2.0,
            '层级': 2,
            '生产周期(天)': 2
        },
        {
            '父物料编码': 'SEMI-Y100',
            '父物料名称': 'Y100半成品',
            '子物料编码': 'RAW-M003',
            '子物料名称': 'M003原材料',
            '用量': 1.5,
            '层级': 2,
            '生产周期(天)': 1
        },
        # PROD-B200的BOM
        {
            '父物料编码': 'PROD-B200',
            '父物料名称': 'B200成品',
            '子物料编码': 'SEMI-Z200',
            '子物料名称': 'Z200半成品',
            '用量': 1.0,
            '层级': 1,
            '生产周期(天)': 4
        },
        {
            '父物料编码': 'PROD-B200',
            '父物料名称': 'B200成品',
            '子物料编码': 'SEMI-Y100',
            '子物料名称': 'Y100半成品',
            '用量': 2.0,
            '层级': 1,
            '生产周期(天)': 3
        },
        # SEMI-Z200的BOM
        {
            '父物料编码': 'SEMI-Z200',
            '父物料名称': 'Z200半成品',
            '子物料编码': 'RAW-M001',
            '子物料名称': 'M001原材料',
            '用量': 2.5,
            '层级': 2,
            '生产周期(天)': 2
        },
        {
            '父物料编码': 'SEMI-Z200',
            '父物料名称': 'Z200半成品',
            '子物料编码': 'RAW-M004',
            '子物料名称': 'M004原材料',
            '用量': 1.0,
            '层级': 2,
            '生产周期(天)': 1
        }
    ]
    
    bom_df = pd.DataFrame(bom_data)
    bom_file = os.path.join(input_dir, 'bom.xlsx')
    bom_df.to_excel(bom_file, index=False)
    print(f"✓ BOM数据已创建: {bom_file}")
    print(f"  - BOM记录数: {len(bom_df)}")
    print(f"  - 最大层级: {bom_df['层级'].max()}")
    
    # ========== 3. 创建产能数据 ==========
    print("\n创建产能数据...")
    
    capacity_data = [
        {
            '物料编码': 'SEMI-X100',
            '物料名称': 'X100半成品',
            '产线/工序': '产线1',
            '日产能上限': 400
        },
        {
            '物料编码': 'SEMI-Y100',
            '物料名称': 'Y100半成品',
            '产线/工序': '产线2',
            '日产能上限': 600
        },
        {
            '物料编码': 'SEMI-Z200',
            '物料名称': 'Z200半成品',
            '产线/工序': '产线3',
            '日产能上限': 300
        },
        {
            '物料编码': 'RAW-M001',
            '物料名称': 'M001原材料',
            '产线/工序': '产线4',
            '日产能上限': 1500
        },
        {
            '物料编码': 'RAW-M002',
            '物料名称': 'M002原材料',
            '产线/工序': '产线5',
            '日产能上限': 1200
        },
        {
            '物料编码': 'RAW-M003',
            '物料名称': 'M003原材料',
            '产线/工序': '产线6',
            '日产能上限': 1000
        },
        {
            '物料编码': 'RAW-M004',
            '物料名称': 'M004原材料',
            '产线/工序': '产线7',
            '日产能上限': 800
        }
    ]
    
    capacity_df = pd.DataFrame(capacity_data)
    capacity_file = os.path.join(input_dir, 'capacity.xlsx')
    capacity_df.to_excel(capacity_file, index=False)
    print(f"✓ 产能数据已创建: {capacity_file}")
    print(f"  - 物料数量: {len(capacity_df)}")
    
    # ========== 完成 ==========
    print("\n" + "=" * 60)
    print("示例数据创建完成！")
    print("=" * 60)
    print(f"\n数据文件位置: {os.path.abspath(input_dir)}")
    print("\n现在可以运行主程序进行测试:")
    print("  python main.py")
    print("\n" + "=" * 60)


if __name__ == "__main__":
    create_sample_data()
