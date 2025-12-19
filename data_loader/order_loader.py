# -*- coding: utf-8 -*-
"""
订单数据加载器
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.validators import DataValidator


class OrderLoader:
    """订单数据加载器"""
    
    # 必填列定义
    REQUIRED_COLUMNS = ['订单号', '产品型号', '数量', '生产开工日期', '发货日期']
    
    # 可选列定义
    OPTIONAL_COLUMNS = ['优先级', '客户名称', '备注']
    
    def __init__(self, file_path):
        """
        初始化订单加载器
        
        Args:
            file_path: 订单Excel文件路径
        """
        self.file_path = file_path
        self.df = None
    
    def load(self):
        """
        加载订单数据
        
        Returns:
            pandas DataFrame: 订单数据
            
        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 数据验证失败
        """
        # 检查文件是否存在
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"订单文件不存在: {self.file_path}")
        
        # 读取Excel文件
        try:
            self.df = pd.read_excel(self.file_path)
        except Exception as e:
            raise ValueError(f"读取订单文件失败: {str(e)}")
        
        # 数据验证
        self._validate()
        
        # 数据转换
        self._transform()
        
        return self.df
    
    def _validate(self):
        """验证数据"""
        # 验证必填列
        DataValidator.validate_columns(self.df, self.REQUIRED_COLUMNS, "订单数据")
        
        # 验证空值
        DataValidator.validate_no_nulls(self.df, self.REQUIRED_COLUMNS, "订单数据")
        
        # 验证订单号唯一性
        DataValidator.validate_unique(self.df, ['订单号'], "订单数据")
        
        # 验证数量为正数
        DataValidator.validate_positive_numbers(self.df, ['数量'], "订单数据")
    
    def _transform(self):
        """数据转换"""
        # 数据类型转换
        type_mapping = {
            '生产开工日期': 'datetime',
            '发货日期': 'datetime',
            '数量': 'int',
            '订单号': 'str',
            '产品型号': 'str'
        }
        
        self.df = DataValidator.validate_data_types(self.df, type_mapping, "订单数据")
        
        # 验证日期逻辑
        DataValidator.validate_date_logic(
            self.df, 
            '生产开工日期', 
            '发货日期', 
            "订单数据"
        )
        
        # 处理可选列
        if '优先级' in self.df.columns:
            self.df['优先级'] = self.df['优先级'].fillna(5).astype(int)
        else:
            self.df['优先级'] = 5  # 默认优先级
        
        # 去除首尾空格
        for col in ['订单号', '产品型号']:
            if col in self.df.columns:
                self.df[col] = self.df[col].str.strip()
    
    def get_summary(self):
        """
        获取订单数据摘要
        
        Returns:
            dict: 摘要信息
        """
        if self.df is None:
            return {}
        
        return {
            '订单总数': len(self.df),
            '产品型号数': self.df['产品型号'].nunique(),
            '总数量': int(self.df['数量'].sum()),
            '最早开工日期': self.df['生产开工日期'].min().strftime('%Y-%m-%d'),
            '最晚交付日期': self.df['发货日期'].max().strftime('%Y-%m-%d')
        }
