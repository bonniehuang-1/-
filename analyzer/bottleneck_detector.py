# -*- coding: utf-8 -*-
"""
瓶颈检测器
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config


class BottleneckDetector:
    """瓶颈检测器"""
    
    def __init__(self, schedule_df, capacity_df):
        """
        初始化瓶颈检测器
        
        Args:
            schedule_df: 排产计划DataFrame
            capacity_df: 产能数据DataFrame
        """
        self.schedule_df = schedule_df
        self.capacity_df = capacity_df
    
    def calculate_gap(self):
        """
        计算产能缺口
        
        Returns:
            pandas DataFrame: 产能缺口明细
        """
        print("  开始计算产能缺口...")
        
        gaps = []
        
        for idx, schedule in self.schedule_df.iterrows():
            material_code = schedule['物料编码']
            total_requirement = schedule['总需求量']
            daily_capacity = schedule['日产能']
            is_delayed = schedule['是否延期']
            delay_days = schedule['延期天数']
            
            if is_delayed and daily_capacity > 0:
                # 计算理论所需天数
                required_days = schedule.get('生产天数', 0)
                
                # 计算可生产量（基于要求完成日期）
                # 这里简化处理，实际可生产量 = 总需求量（因为已经按实际产能排产）
                actual_production = total_requirement
                
                # 缺口数量 = 延期天数 * 日产能（理论上需要额外的产能）
                gap_quantity = delay_days * daily_capacity
                gap_rate = (gap_quantity / total_requirement * 100) if total_requirement > 0 else 0
                
                gaps.append({
                    '物料编码': material_code,
                    '物料名称': schedule.get('物料名称', ''),
                    '总需求量': total_requirement,
                    '日产能': daily_capacity,
                    '可生产量': actual_production,
                    '缺口数量': gap_quantity,
                    '缺口率(%)': round(gap_rate, 2),
                    '延期天数': delay_days,
                    '平均产能利用率(%)': round(schedule.get('平均产能利用率', 0) * 100, 2)
                })
        
        gap_df = pd.DataFrame(gaps)
        
        if not gap_df.empty:
            gap_df = gap_df.sort_values('缺口数量', ascending=False)
        
        print(f"  完成产能缺口计算，发现{len(gap_df)}个瓶颈物料")
        
        return gap_df
    
    def summarize(self):
        """
        汇总瓶颈信息
        
        Returns:
            pandas DataFrame: 瓶颈物料汇总
        """
        print("  开始汇总瓶颈信息...")
        
        bottlenecks = []
        
        for idx, schedule in self.schedule_df.iterrows():
            material_code = schedule['物料编码']
            avg_utilization = schedule.get('平均产能利用率', 0)
            is_delayed = schedule['是否延期']
            
            # 判断是否为瓶颈
            is_bottleneck = False
            bottleneck_type = '正常'
            
            if is_delayed:
                is_bottleneck = True
                bottleneck_type = '产能不足-延期'
            elif avg_utilization >= Config.CAPACITY_UTILIZATION_THRESHOLD:
                is_bottleneck = True
                bottleneck_type = '产能紧张'
            
            if is_bottleneck:
                bottlenecks.append({
                    '物料编码': material_code,
                    '物料名称': schedule.get('物料名称', ''),
                    '瓶颈类型': bottleneck_type,
                    '日产能': schedule['日产能'],
                    '总需求量': schedule['总需求量'],
                    '产能利用率(%)': round(avg_utilization * 100, 2),
                    '延期天数': schedule['延期天数'],
                    '影响程度': self._calculate_impact(schedule)
                })
        
        bottleneck_df = pd.DataFrame(bottlenecks)
        
        if not bottleneck_df.empty:
            bottleneck_df = bottleneck_df.sort_values('影响程度', ascending=False)
        
        print(f"  完成瓶颈汇总，发现{len(bottleneck_df)}个瓶颈物料")
        
        return bottleneck_df
    
    def _calculate_impact(self, schedule):
        """
        计算影响程度（用于排序）
        
        Args:
            schedule: 排产计划记录
            
        Returns:
            float: 影响程度分数
        """
        # 影响程度 = 延期天数 * 10 + 产能利用率 * 100
        delay_days = schedule.get('延期天数', 0)
        utilization = schedule.get('平均产能利用率', 0)
        
        impact = delay_days * 10 + utilization * 100
        
        return round(impact, 2)
    
    def get_top_bottlenecks(self, bottleneck_df, top_n=10):
        """
        获取TOP N瓶颈物料
        
        Args:
            bottleneck_df: 瓶颈汇总DataFrame
            top_n: 返回前N个
            
        Returns:
            pandas DataFrame: TOP N瓶颈物料
        """
        if bottleneck_df.empty:
            return bottleneck_df
        
        return bottleneck_df.head(top_n)
    
    def get_capacity_recommendations(self, gap_df):
        """
        生成产能提升建议
        
        Args:
            gap_df: 产能缺口DataFrame
            
        Returns:
            list: 建议列表
        """
        recommendations = []
        
        for idx, gap in gap_df.iterrows():
            material_code = gap['物料编码']
            gap_quantity = gap['缺口数量']
            daily_capacity = gap['日产能']
            delay_days = gap['延期天数']
            
            # 计算需要提升的产能
            required_capacity_increase = gap_quantity / delay_days if delay_days > 0 else 0
            increase_rate = (required_capacity_increase / daily_capacity * 100) if daily_capacity > 0 else 0
            
            recommendation = {
                '物料编码': material_code,
                '物料名称': gap.get('物料名称', ''),
                '当前日产能': daily_capacity,
                '建议日产能': int(daily_capacity + required_capacity_increase),
                '需提升产能': int(required_capacity_increase),
                '提升比例(%)': round(increase_rate, 2),
                '预期效果': f'可消除{delay_days}天延期'
            }
            
            recommendations.append(recommendation)
        
        return recommendations
