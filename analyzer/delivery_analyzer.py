# -*- coding: utf-8 -*-
"""
交付能力分析器
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import Config


class DeliveryAnalyzer:
    """交付能力分析器"""
    
    def __init__(self, orders_df, schedule_df, bom_df, mrp_df):
        """
        初始化交付分析器
        
        Args:
            orders_df: 订单数据DataFrame
            schedule_df: 排产计划DataFrame
            bom_df: BOM数据DataFrame
            mrp_df: 物料需求计划DataFrame
        """
        self.orders_df = orders_df
        self.schedule_df = schedule_df
        self.bom_df = bom_df
        self.mrp_df = mrp_df
    
    def analyze(self):
        """
        分析交付能力
        
        Returns:
            pandas DataFrame: 交付能力分析结果
        """
        print("  开始分析交付能力...")
        
        results = []
        
        for idx, order in self.orders_df.iterrows():
            order_no = order['订单号']
            product_code = order['产品型号']
            quantity = order['数量']
            required_delivery = order['发货日期']
            start_date = order['生产开工日期']
            
            # 找出该订单涉及的所有物料
            materials = self._get_order_materials(product_code)
            
            # 找出关键路径（最晚完成的物料）
            bottleneck_material, estimated_finish = self._find_critical_path(
                materials, 
                order_no
            )
            
            # 判断能否按时交付
            if estimated_finish is None:
                # 无法确定完成时间（可能缺少产能数据）
                can_deliver = False
                delay_days = 9999
                status = '无法排产'
            else:
                can_deliver = estimated_finish <= required_delivery
                if can_deliver:
                    delay_days = 0
                    status = '正常'
                else:
                    delay_days = (estimated_finish - required_delivery).days
                    status = self._get_status(delay_days)
            
            results.append({
                '订单号': order_no,
                '产品型号': product_code,
                '数量': quantity,
                '生产开工日期': start_date,
                '要求交付日期': required_delivery,
                '预计完成日期': estimated_finish,
                '能否按时交付': can_deliver,
                '延期天数': delay_days,
                '瓶颈物料': bottleneck_material if bottleneck_material else '未知',
                '状态': status,
                '涉及物料数': len(materials)
            })
        
        analysis_df = pd.DataFrame(results)
        
        print(f"  完成交付能力分析，共{len(analysis_df)}个订单")
        
        return analysis_df
    
    def _get_order_materials(self, product_code):
        """
        获取订单涉及的所有物料
        
        Args:
            product_code: 产品编码
            
        Returns:
            list: 物料编码列表
        """
        materials = set()
        
        def traverse_bom(material):
            """递归遍历BOM"""
            children = self.bom_df[self.bom_df['父物料编码'] == material]
            for _, child in children.iterrows():
                child_code = child['子物料编码']
                materials.add(child_code)
                traverse_bom(child_code)
        
        traverse_bom(product_code)
        return list(materials)
    
    def _find_critical_path(self, materials, order_no):
        """
        找出关键路径（最晚完成的物料）
        
        Args:
            materials: 物料编码列表
            order_no: 订单号
            
        Returns:
            tuple: (瓶颈物料编码, 最晚完成日期)
        """
        latest_finish = None
        bottleneck = None
        
        for material in materials:
            # 从排产计划中查找该物料
            schedule = self.schedule_df[
                self.schedule_df['物料编码'] == material
            ]
            
            if not schedule.empty:
                finish_date_str = schedule.iloc[0]['预计完成日期']
                
                if finish_date_str and finish_date_str != 'None':
                    finish_date = pd.to_datetime(finish_date_str)
                    
                    if latest_finish is None or finish_date > latest_finish:
                        latest_finish = finish_date
                        bottleneck = material
        
        return bottleneck, latest_finish
    
    def _get_status(self, delay_days):
        """
        获取状态标识
        
        Args:
            delay_days: 延期天数
            
        Returns:
            str: 状态标识
        """
        if delay_days == 0:
            return '正常'
        elif delay_days >= Config.ALERT_DELAY_DAYS_RED:
            return '红色预警'
        elif delay_days >= Config.ALERT_DELAY_DAYS_YELLOW:
            return '黄色预警'
        else:
            return '正常'
    
    def get_summary(self, analysis_df):
        """
        获取分析摘要
        
        Args:
            analysis_df: 分析结果DataFrame
            
        Returns:
            dict: 摘要信息
        """
        total_orders = len(analysis_df)
        on_time_orders = len(analysis_df[analysis_df['能否按时交付']])
        delayed_orders = total_orders - on_time_orders
        
        red_alerts = len(analysis_df[analysis_df['状态'] == '红色预警'])
        yellow_alerts = len(analysis_df[analysis_df['状态'] == '黄色预警'])
        
        total_delay_days = analysis_df[~analysis_df['能否按时交付']]['延期天数'].sum()
        
        return {
            '订单总数': total_orders,
            '按时交付订单数': on_time_orders,
            '延期订单数': delayed_orders,
            '按时交付率': round(on_time_orders / total_orders * 100, 2) if total_orders > 0 else 0,
            '红色预警数': red_alerts,
            '黄色预警数': yellow_alerts,
            '总延期天数': int(total_delay_days) if not pd.isna(total_delay_days) else 0
        }
    
    def get_delayed_orders(self, analysis_df):
        """
        获取延期订单列表
        
        Args:
            analysis_df: 分析结果DataFrame
            
        Returns:
            pandas DataFrame: 延期订单
        """
        return analysis_df[~analysis_df['能否按时交付']].sort_values(
            '延期天数', 
            ascending=False
        )
    
    def get_alerts(self, analysis_df):
        """
        获取预警信息
        
        Args:
            analysis_df: 分析结果DataFrame
            
        Returns:
            dict: 预警信息字典
        """
        red_alerts = analysis_df[analysis_df['状态'] == '红色预警']
        yellow_alerts = analysis_df[analysis_df['状态'] == '黄色预警']
        
        return {
            'red': red_alerts.to_dict('records'),
            'yellow': yellow_alerts.to_dict('records')
        }
