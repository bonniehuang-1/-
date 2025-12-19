# -*- coding: utf-8 -*-
"""
物料需求计算器 (MRP - Material Requirement Planning)
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.date_utils import DateUtils


class MaterialRequirementCalculator:
    """物料需求计划计算器"""
    
    def __init__(self, orders_df, bom_df):
        """
        初始化MRP计算器
        
        Args:
            orders_df: 订单数据DataFrame
            bom_df: BOM数据DataFrame
        """
        self.orders_df = orders_df
        self.bom_df = bom_df
        self.requirements = {}  # 物料需求字典
    
    def calculate(self):
        """
        计算物料需求计划
        
        Returns:
            pandas DataFrame: 物料需求计划
        """
        print("  开始计算物料需求...")
        
        # 遍历所有订单
        for idx, order in self.orders_df.iterrows():
            order_no = order['订单号']
            product_code = order['产品型号']
            quantity = order['数量']
            delivery_date = order['发货日期']
            
            # 递归展开BOM
            self._explode_bom(
                material_code=product_code,
                quantity=quantity,
                required_date=delivery_date,
                order_no=order_no,
                level=0
            )
        
        # 转换为DataFrame
        mrp_df = self._convert_to_dataframe()
        
        print(f"  完成物料需求计算，共{len(mrp_df)}个物料")
        
        return mrp_df
    
    def _explode_bom(self, material_code, quantity, required_date, order_no, level):
        """
        递归展开BOM
        
        Args:
            material_code: 物料编码
            quantity: 需求数量
            required_date: 要求完成日期
            order_no: 订单号
            level: 当前层级
        """
        # 查找该物料的子物料
        children = self.bom_df[self.bom_df['父物料编码'] == material_code]
        
        if children.empty:
            # 叶子节点（原材料或外购件），不需要进一步展开
            return
        
        # 遍历所有子物料
        for _, child in children.iterrows():
            child_code = child['子物料编码']
            usage = child['用量']
            lead_time = child['生产周期(天)']
            child_level = child['层级']
            
            # 计算子物料需求量
            child_quantity = quantity * usage
            
            # 计算子物料最晚完成日期（考虑生产周期）
            child_required_date = DateUtils.subtract_workdays(required_date, lead_time)
            
            # 记录需求
            if child_code not in self.requirements:
                self.requirements[child_code] = {
                    '物料编码': child_code,
                    '物料名称': child.get('子物料名称', ''),
                    '层级': child_level,
                    '生产周期': lead_time,
                    '总需求量': 0,
                    '需求明细': []
                }
            
            # 累加需求量
            self.requirements[child_code]['总需求量'] += child_quantity
            
            # 添加需求明细
            self.requirements[child_code]['需求明细'].append({
                '订单号': order_no,
                '需求量': child_quantity,
                '最晚完成日期': child_required_date,
                '父物料': material_code
            })
            
            # 递归处理下一层
            self._explode_bom(
                material_code=child_code,
                quantity=child_quantity,
                required_date=child_required_date,
                order_no=order_no,
                level=level + 1
            )
    
    def _convert_to_dataframe(self):
        """
        将需求字典转换为DataFrame
        
        Returns:
            pandas DataFrame: 物料需求计划
        """
        if not self.requirements:
            return pd.DataFrame()
        
        # 准备数据列表
        data = []
        for material_code, req in self.requirements.items():
            # 找出最早的要求完成日期
            earliest_required_date = min(
                detail['最晚完成日期'] for detail in req['需求明细']
            )
            
            # 找出最晚的要求完成日期
            latest_required_date = max(
                detail['最晚完成日期'] for detail in req['需求明细']
            )
            
            # 统计涉及的订单数
            order_count = len(set(detail['订单号'] for detail in req['需求明细']))
            
            data.append({
                '物料编码': material_code,
                '物料名称': req['物料名称'],
                '层级': req['层级'],
                '生产周期': req['生产周期'],
                '总需求量': req['总需求量'],
                '最早要求日期': earliest_required_date,
                '最晚要求日期': latest_required_date,
                '涉及订单数': order_count,
                '需求明细': req['需求明细']
            })
        
        # 创建DataFrame并按层级和最早要求日期排序
        df = pd.DataFrame(data)
        df = df.sort_values(['层级', '最早要求日期'], ascending=[True, True])
        df = df.reset_index(drop=True)
        
        return df
    
    def get_material_requirement(self, material_code):
        """
        获取指定物料的需求信息
        
        Args:
            material_code: 物料编码
            
        Returns:
            dict: 需求信息，如果不存在返回None
        """
        return self.requirements.get(material_code)
    
    def get_summary(self):
        """
        获取MRP计算摘要
        
        Returns:
            dict: 摘要信息
        """
        if not self.requirements:
            return {}
        
        total_materials = len(self.requirements)
        total_quantity = sum(req['总需求量'] for req in self.requirements.values())
        
        # 按层级统计
        level_stats = {}
        for req in self.requirements.values():
            level = req['层级']
            if level not in level_stats:
                level_stats[level] = 0
            level_stats[level] += 1
        
        return {
            '物料总数': total_materials,
            '总需求量': round(total_quantity, 2),
            '层级分布': level_stats
        }
