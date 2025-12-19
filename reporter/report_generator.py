# -*- coding: utf-8 -*-
"""
报告生成器
"""
import pandas as pd
import xlsxwriter
from datetime import datetime
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class ReportGenerator:
    """Excel报告生成器"""
    
    def __init__(self, output_path):
        """
        初始化报告生成器
        
        Args:
            output_path: 输出文件路径
        """
        self.output_path = output_path
        self.workbook = None
        self.formats = {}
    
    def generate(self, delivery_analysis, capacity_gap, bottleneck_summary, 
                 summary_stats=None):
        """
        生成完整的Excel报告
        
        Args:
            delivery_analysis: 交付能力分析DataFrame
            capacity_gap: 产能缺口DataFrame
            bottleneck_summary: 瓶颈汇总DataFrame
            summary_stats: 汇总统计信息字典
        """
        print("  开始生成Excel报告...")
        
        # 创建Excel文件
        self.workbook = xlsxwriter.Workbook(self.output_path)
        
        # 创建格式
        self._create_formats()
        
        # 生成各个工作表
        self._create_summary_sheet(delivery_analysis, summary_stats)
        self._create_delivery_detail_sheet(delivery_analysis)
        self._create_capacity_gap_sheet(capacity_gap)
        self._create_bottleneck_sheet(bottleneck_summary)
        
        # 关闭文件
        self.workbook.close()
        
        print(f"  报告已生成: {self.output_path}")
    
    def _create_formats(self):
        """创建单元格格式"""
        self.formats = {
            'title': self.workbook.add_format({
                'bold': True,
                'font_size': 14,
                'align': 'center',
                'valign': 'vcenter',
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1
            }),
            'header': self.workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            }),
            'red_alert': self.workbook.add_format({
                'bg_color': '#FF0000',
                'font_color': 'white',
                'border': 1
            }),
            'yellow_alert': self.workbook.add_format({
                'bg_color': '#FFFF00',
                'border': 1
            }),
            'green_normal': self.workbook.add_format({
                'bg_color': '#00FF00',
                'border': 1
            }),
            'normal': self.workbook.add_format({
                'border': 1
            }),
            'number': self.workbook.add_format({
                'border': 1,
                'num_format': '#,##0'
            }),
            'percent': self.workbook.add_format({
                'border': 1,
                'num_format': '0.00%'
            }),
            'date': self.workbook.add_format({
                'border': 1,
                'num_format': 'yyyy-mm-dd'
            })
        }
    
    def _create_summary_sheet(self, delivery_df, summary_stats):
        """创建汇总表"""
        worksheet = self.workbook.add_worksheet('交付能力总览')
        
        # 设置列宽
        worksheet.set_column('A:A', 15)
        worksheet.set_column('B:K', 12)
        
        row = 0
        
        # 标题
        worksheet.merge_range(row, 0, row, 7, '生产排产交付能力分析报告', self.formats['title'])
        worksheet.set_row(row, 25)
        row += 1
        
        # 生成时间
        worksheet.write(row, 0, '生成时间:', self.formats['normal'])
        worksheet.write(row, 1, datetime.now().strftime('%Y-%m-%d %H:%M:%S'), self.formats['normal'])
        row += 2
        
        # 汇总统计
        if summary_stats:
            worksheet.write(row, 0, '汇总统计', self.formats['header'])
            row += 1
            
            for key, value in summary_stats.items():
                worksheet.write(row, 0, key, self.formats['normal'])
                if isinstance(value, (int, float)):
                    worksheet.write(row, 1, value, self.formats['number'])
                else:
                    worksheet.write(row, 1, value, self.formats['normal'])
                row += 1
            
            row += 1
        
        # 订单明细表头
        headers = ['订单号', '产品型号', '数量', '开工日期', '要求交付日期', 
                   '预计完成日期', '延期天数', '状态', '瓶颈物料']
        
        for col, header in enumerate(headers):
            worksheet.write(row, col, header, self.formats['header'])
        row += 1
        
        # 写入数据
        for _, record in delivery_df.iterrows():
            worksheet.write(row, 0, str(record['订单号']), self.formats['normal'])
            worksheet.write(row, 1, str(record['产品型号']), self.formats['normal'])
            worksheet.write(row, 2, int(record['数量']), self.formats['number'])
            
            # 日期处理
            if pd.notna(record['生产开工日期']):
                worksheet.write(row, 3, record['生产开工日期'].strftime('%Y-%m-%d'), self.formats['normal'])
            
            if pd.notna(record['要求交付日期']):
                worksheet.write(row, 4, record['要求交付日期'].strftime('%Y-%m-%d'), self.formats['normal'])
            
            if pd.notna(record['预计完成日期']):
                worksheet.write(row, 5, record['预计完成日期'].strftime('%Y-%m-%d'), self.formats['normal'])
            else:
                worksheet.write(row, 5, '无法确定', self.formats['normal'])
            
            worksheet.write(row, 6, int(record['延期天数']), self.formats['number'])
            
            # 根据状态设置格式
            status = record['状态']
            if status == '红色预警':
                fmt = self.formats['red_alert']
            elif status == '黄色预警':
                fmt = self.formats['yellow_alert']
            elif status == '正常':
                fmt = self.formats['green_normal']
            else:
                fmt = self.formats['normal']
            
            worksheet.write(row, 7, status, fmt)
            worksheet.write(row, 8, str(record['瓶颈物料']), self.formats['normal'])
            
            row += 1
    
    def _create_delivery_detail_sheet(self, delivery_df):
        """创建交付详情表"""
        worksheet = self.workbook.add_worksheet('延期订单明细')
        
        # 筛选延期订单
        delayed_df = delivery_df[~delivery_df['能否按时交付']].copy()
        delayed_df = delayed_df.sort_values('延期天数', ascending=False)
        
        # 写入表头
        headers = ['订单号', '产品型号', '数量', '要求交付日期', '预计完成日期', 
                   '延期天数', '状态', '瓶颈物料', '涉及物料数']
        
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, self.formats['header'])
        
        # 写入数据
        for idx, (_, record) in enumerate(delayed_df.iterrows(), start=1):
            worksheet.write(idx, 0, str(record['订单号']), self.formats['normal'])
            worksheet.write(idx, 1, str(record['产品型号']), self.formats['normal'])
            worksheet.write(idx, 2, int(record['数量']), self.formats['number'])
            
            if pd.notna(record['要求交付日期']):
                worksheet.write(idx, 3, record['要求交付日期'].strftime('%Y-%m-%d'), self.formats['normal'])
            
            if pd.notna(record['预计完成日期']):
                worksheet.write(idx, 4, record['预计完成日期'].strftime('%Y-%m-%d'), self.formats['normal'])
            
            worksheet.write(idx, 5, int(record['延期天数']), self.formats['number'])
            
            status = record['状态']
            fmt = self.formats['red_alert'] if status == '红色预警' else self.formats['yellow_alert']
            worksheet.write(idx, 6, status, fmt)
            
            worksheet.write(idx, 7, str(record['瓶颈物料']), self.formats['normal'])
            worksheet.write(idx, 8, int(record['涉及物料数']), self.formats['number'])
        
        # 设置列宽
        worksheet.set_column('A:I', 15)
    
    def _create_capacity_gap_sheet(self, gap_df):
        """创建产能缺口表"""
        worksheet = self.workbook.add_worksheet('产能缺口明细')
        
        # 写入表头
        headers = ['物料编码', '物料名称', '总需求量', '日产能', '缺口数量', 
                   '缺口率(%)', '延期天数', '产能利用率(%)']
        
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, self.formats['header'])
        
        # 写入数据
        for idx, (_, record) in enumerate(gap_df.iterrows(), start=1):
            worksheet.write(idx, 0, str(record['物料编码']), self.formats['normal'])
            worksheet.write(idx, 1, str(record.get('物料名称', '')), self.formats['normal'])
            worksheet.write(idx, 2, record['总需求量'], self.formats['number'])
            worksheet.write(idx, 3, int(record['日产能']), self.formats['number'])
            worksheet.write(idx, 4, record['缺口数量'], self.formats['number'])
            worksheet.write(idx, 5, record['缺口率(%)'], self.formats['number'])
            worksheet.write(idx, 6, int(record['延期天数']), self.formats['number'])
            worksheet.write(idx, 7, record['平均产能利用率(%)'], self.formats['number'])
        
        # 设置列宽
        worksheet.set_column('A:H', 15)
    
    def _create_bottleneck_sheet(self, bottleneck_df):
        """创建瓶颈分析表"""
        worksheet = self.workbook.add_worksheet('瓶颈物料分析')
        
        # 写入表头
        headers = ['物料编码', '物料名称', '瓶颈类型', '日产能', '总需求量', 
                   '产能利用率(%)', '延期天数', '影响程度']
        
        for col, header in enumerate(headers):
            worksheet.write(0, col, header, self.formats['header'])
        
        # 写入数据
        for idx, (_, record) in enumerate(bottleneck_df.iterrows(), start=1):
            worksheet.write(idx, 0, str(record['物料编码']), self.formats['normal'])
            worksheet.write(idx, 1, str(record.get('物料名称', '')), self.formats['normal'])
            worksheet.write(idx, 2, str(record['瓶颈类型']), self.formats['normal'])
            worksheet.write(idx, 3, int(record['日产能']), self.formats['number'])
            worksheet.write(idx, 4, record['总需求量'], self.formats['number'])
            worksheet.write(idx, 5, record['产能利用率(%)'], self.formats['number'])
            worksheet.write(idx, 6, int(record['延期天数']), self.formats['number'])
            worksheet.write(idx, 7, record['影响程度'], self.formats['number'])
        
        # 设置列宽
        worksheet.set_column('A:H', 15)
