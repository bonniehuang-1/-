# -*- coding: utf-8 -*-
"""
产能数据加载器
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.validators import DataValidator


class CapacityLoader:
    """产能数据加载器"""
    
    # 必填列定义
    REQUIRED_COLUMNS = ['物料编码', '日产能上限']
    
    # 可选列定义
    OPTIONAL_COLUMNS = ['物料名称', '产线/工序', '生效日期', '备注']
    
    def __init__(self, file_path):
        """
        初始化产能加载器
        
        Args:
            file_path: 产能Excel文件路径
        """
        self.file_path = file_path
        self.df = None
    
    def load(self):
        """
        加载产能数据
        
        Returns:
            pandas DataFrame: 产能数据
            
        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 数据验证失败
        """
        # 检查文件是否存在
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"产能文件不存在: {self.file_path}")
        
        # 读取Excel文件
        try:
            self.df = pd.read_excel(self.file_path)
        except Exception as e:
            raise ValueError(f"读取产能文件失败: {str(e)}")
        
        # 数据验证
        self._validate()
        
        # 数据转换
        self._transform()
        
        return self.df
    
    def _validate(self):
        """验证数据"""
        # 验证必填列
        DataValidator.validate_columns(self.df, self.REQUIRED_COLUMNS, "产能数据")
        
        # 验证空值
        DataValidator.validate_no_nulls(self.df, self.REQUIRED_COLUMNS, "产能数据")
        
        # 验证产能为正数
        DataValidator.validate_positive_numbers(self.df, ['日产能上限'], "产能数据")
    
    def _transform(self):
        """数据转换"""
        # 数据类型转换
        type_mapping = {
            '物料编码': 'str',
            '日产能上限': 'int'
        }
        
        self.df = DataValidator.validate_data_types(self.df, type_mapping, "产能数据")
        
        # 去除首尾空格
        if '物料编码' in self.df.columns:
            self.df['物料编码'] = self.df['物料编码'].str.strip()
        
        # 处理可选列
        if '产线/工序' not in self.df.columns:
            self.df['产线/工序'] = '默认产线'
        
        # 处理生效日期
        if '生效日期' in self.df.columns:
            self.df['生效日期'] = pd.to_datetime(self.df['生效日期'], errors='coerce')
        
        # 检查物料编码重复（同一物料可能有多个产线，这里简化处理，取第一个）
        if self.df['物料编码'].duplicated().any():
            print("警告: 产能数据中存在重复的物料编码，将保留第一条记录")
            self.df = self.df.drop_duplicates(subset=['物料编码'], keep='first')
    
    def get_summary(self):
        """
        获取产能数据摘要
        
        Returns:
            dict: 摘要信息
        """
        if self.df is None:
            return {}
        
        return {
            '物料数量': len(self.df),
            '总日产能': int(self.df['日产能上限'].sum()),
            '平均日产能': round(self.df['日产能上限'].mean(), 2),
            '最大日产能': int(self.df['日产能上限'].max()),
            '最小日产能': int(self.df['日产能上限'].min())
        }
    
    def get_capacity(self, material_code):
        """
        获取指定物料的产能
        
        Args:
            material_code: 物料编码
            
        Returns:
            int: 日产能上限，如果不存在返回None
        """
        if self.df is None:
            return None
        
        result = self.df[self.df['物料编码'] == material_code]
        if result.empty:
            return None
        
        return int(result.iloc[0]['日产能上限'])
    
    def has_capacity(self, material_code):
        """
        检查指定物料是否有产能数据
        
        Args:
            material_code: 物料编码
            
        Returns:
            bool: True表示有产能数据，False表示没有
        """
        if self.df is None:
            return False
        
        return material_code in self.df['物料编码'].values
