# -*- coding: utf-8 -*-
"""
生产排产调度器
"""
import pandas as pd
from datetime import timedelta
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.date_utils import DateUtils


class ProductionScheduler:
    """生产排产调度器"""
    
    def __init__(self, mrp_df, capacity_df, start_date):
        """
        初始化排产调度器
        
        Args:
            mrp_df: 物料需求计划DataFrame
            capacity_df: 产能数据DataFrame
            start_date: 生产开工日期
        """
        self.mrp_df = mrp_df
        self.capacity_df = capacity_df
        self.start_date = start_date
        self.schedule_results = []
    
    def schedule(self):
        """
        执行排产计算
        
        Returns:
            pandas DataFrame: 排产计划
        """
        print("  开始执行产能排产...")
        
        # 按物料分组排产
        for idx, material_row in self.mrp_df.iterrows():
            material_code = material_row['物料编码']
            
            # 获取该物料的产能信息
            capacity_info = self.capacity_df[
                self.capacity_df['物料编码'] == material_code
            ]
            
            if capacity_info.empty:
                # 无产能数据，记录警告
                print(f"  警告: 物料 {material_code} 没有产能数据，跳过排产")
                self.schedule_results.append({
                    '物料编码': material_code,
                    '物料名称': material_row.get('物料名称', ''),
                    '总需求量': material_row['总需求量'],
                    '日产能': 0,
                    '开工日期': None,
                    '预计完成日期': None,
                    '要求完成日期': material_row['最早要求日期'],
                    '是否延期': True,
                    '延期天数': 9999,
                    '产能状态': '无产能数据',
                    '排产明细': []
                })
                continue
            
            daily_capacity = int(capacity_info.iloc[0]['日产能上限'])
            
            # 计算该物料的排产计划
            schedule = self._schedule_material(
                material_code=material_code,
                material_name=material_row.get('物料名称', ''),
                total_requirement=material_row['总需求量'],
                required_date=material_row['最早要求日期'],
                daily_capacity=daily_capacity
            )
            
            self.schedule_results.append(schedule)
        
        # 转换为DataFrame
        schedule_df = pd.DataFrame(self.schedule_results)
        
        print(f"  完成产能排产，共{len(schedule_df)}个物料")
        
        return schedule_df
    
    def _schedule_material(self, material_code, material_name, total_requirement, 
                          required_date, daily_capacity):
        """
        为单个物料排产
        
        Args:
            material_code: 物料编码
            material_name: 物料名称
            total_requirement: 总需求量
            required_date: 要求完成日期
            daily_capacity: 日产能上限
            
        Returns:
            dict: 排产结果
        """
        # 从开工日期开始排产
        current_date = self.start_date
        cumulative_production = 0
        daily_schedule = []
        days_count = 0
        
        while cumulative_production < total_requirement:
            # 检查是否为工作日
            if DateUtils.is_workday(current_date):
                # 计算当日生产量
                remaining = total_requirement - cumulative_production
                daily_production = min(daily_capacity, remaining)
                
                cumulative_production += daily_production
                days_count += 1
                
                daily_schedule.append({
                    '日期': DateUtils.format_date(current_date),
                    '计划产量': daily_production,
                    '累计产量': cumulative_production,
                    '剩余需求': total_requirement - cumulative_production,
                    '产能利用率': round(daily_production / daily_capacity, 4)
                })
            
            current_date = current_date + timedelta(days=1)
        
        # 预计完成日期（最后一个生产日）
        estimated_finish_date = current_date - timedelta(days=1)
        
        # 判断是否延期
        is_delayed = estimated_finish_date > required_date
        
        if is_delayed:
            delay_days = (estimated_finish_date - required_date).days
            status = '延期'
        else:
            delay_days = 0
            status = '正常'
        
        # 计算平均产能利用率
        if daily_schedule:
            avg_utilization = sum(d['产能利用率'] for d in daily_schedule) / len(daily_schedule)
        else:
            avg_utilization = 0
        
        return {
            '物料编码': material_code,
            '物料名称': material_name,
            '总需求量': total_requirement,
            '日产能': daily_capacity,
            '开工日期': DateUtils.format_date(self.start_date),
            '预计完成日期': DateUtils.format_date(estimated_finish_date),
            '要求完成日期': DateUtils.format_date(required_date),
            '是否延期': is_delayed,
            '延期天数': delay_days,
            '生产天数': days_count,
            '平均产能利用率': round(avg_utilization, 4),
            '产能状态': status,
            '排产明细': daily_schedule
        }
    
    def get_summary(self):
        """
        获取排产摘要
        
        Returns:
            dict: 摘要信息
        """
        if not self.schedule_results:
            return {}
        
        total_materials = len(self.schedule_results)
        delayed_materials = sum(1 for s in self.schedule_results if s['是否延期'])
        
        # 计算总延期天数
        total_delay_days = sum(s['延期天数'] for s in self.schedule_results if s['是否延期'])
        
        # 计算平均产能利用率
        valid_schedules = [s for s in self.schedule_results if s.get('平均产能利用率', 0) > 0]
        if valid_schedules:
            avg_utilization = sum(s['平均产能利用率'] for s in valid_schedules) / len(valid_schedules)
        else:
            avg_utilization = 0
        
        return {
            '物料总数': total_materials,
            '正常物料数': total_materials - delayed_materials,
            '延期物料数': delayed_materials,
            '延期率': round(delayed_materials / total_materials * 100, 2) if total_materials > 0 else 0,
            '总延期天数': total_delay_days,
            '平均产能利用率': round(avg_utilization * 100, 2)
        }
    
    def get_material_schedule(self, material_code):
        """
        获取指定物料的排产计划
        
        Args:
            material_code: 物料编码
            
        Returns:
            dict: 排产计划，如果不存在返回None
        """
        for schedule in self.schedule_results:
            if schedule['物料编码'] == material_code:
                return schedule
        return None
