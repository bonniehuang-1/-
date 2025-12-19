# -*- coding: utf-8 -*-
"""
BOM物料清单数据加载器
"""
import pandas as pd
import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.validators import DataValidator
from config import Config


class BOMLoader:
    """BOM物料清单数据加载器"""
    
    # 必填列定义
    REQUIRED_COLUMNS = ['父物料编码', '子物料编码', '用量', '层级', '生产周期(天)']
    
    # 可选列定义
    OPTIONAL_COLUMNS = ['父物料名称', '子物料名称', '单位', '备注']
    
    def __init__(self, file_path):
        """
        初始化BOM加载器
        
        Args:
            file_path: BOM Excel文件路径
        """
        self.file_path = file_path
        self.df = None
    
    def load(self):
        """
        加载BOM数据
        
        Returns:
            pandas DataFrame: BOM数据
            
        Raises:
            FileNotFoundError: 文件不存在
            ValueError: 数据验证失败
        """
        # 检查文件是否存在
        if not os.path.exists(self.file_path):
            raise FileNotFoundError(f"BOM文件不存在: {self.file_path}")
        
        # 读取Excel文件
        try:
            self.df = pd.read_excel(self.file_path)
        except Exception as e:
            raise ValueError(f"读取BOM文件失败: {str(e)}")
        
        # 数据验证
        self._validate()
        
        # 数据转换
        self._transform()
        
        # 检测循环引用
        self._detect_circular_reference()
        
        return self.df
    
    def _validate(self):
        """验证数据"""
        # 验证必填列
        DataValidator.validate_columns(self.df, self.REQUIRED_COLUMNS, "BOM数据")
        
        # 验证空值
        DataValidator.validate_no_nulls(self.df, self.REQUIRED_COLUMNS, "BOM数据")
        
        # 验证数量为正数
        DataValidator.validate_positive_numbers(
            self.df, 
            ['用量', '生产周期(天)'], 
            "BOM数据"
        )
    
    def _transform(self):
        """数据转换"""
        # 数据类型转换
        type_mapping = {
            '父物料编码': 'str',
            '子物料编码': 'str',
            '用量': 'float',
            '层级': 'int',
            '生产周期(天)': 'int'
        }
        
        self.df = DataValidator.validate_data_types(self.df, type_mapping, "BOM数据")
        
        # 去除首尾空格
        for col in ['父物料编码', '子物料编码']:
            if col in self.df.columns:
                self.df[col] = self.df[col].str.strip()
        
        # 验证层级合理性
        if (self.df['层级'] > Config.MAX_BOM_LEVEL).any():
            raise ValueError(f"BOM层级超过最大限制{Config.MAX_BOM_LEVEL}")
        
        if (self.df['层级'] < 1).any():
            raise ValueError("BOM层级必须大于等于1")
    
    def _detect_circular_reference(self):
        """
        检测BOM循环引用
        
        Raises:
            ValueError: 如果存在循环引用
        """
        # 构建物料关系图
        graph = {}
        for _, row in self.df.iterrows():
            parent = row['父物料编码']
            child = row['子物料编码']
            
            # 检查父子物料是否相同
            if parent == child:
                raise ValueError(f"BOM数据存在自引用: {parent}")
            
            if parent not in graph:
                graph[parent] = []
            graph[parent].append(child)
        
        # 使用DFS检测环
        def has_cycle(node, visited, rec_stack, path):
            """深度优先搜索检测环"""
            visited.add(node)
            rec_stack.add(node)
            path.append(node)
            
            if node in graph:
                for neighbor in graph[node]:
                    if neighbor not in visited:
                        if has_cycle(neighbor, visited, rec_stack, path):
                            return True
                    elif neighbor in rec_stack:
                        # 找到环，构建环路径
                        cycle_start = path.index(neighbor)
                        cycle_path = ' -> '.join(path[cycle_start:] + [neighbor])
                        raise ValueError(f"BOM数据存在循环引用: {cycle_path}")
            
            path.pop()
            rec_stack.remove(node)
            return False
        
        visited = set()
        for node in graph:
            if node not in visited:
                has_cycle(node, visited, set(), [])
    
    def get_summary(self):
        """
        获取BOM数据摘要
        
        Returns:
            dict: 摘要信息
        """
        if self.df is None:
            return {}
        
        return {
            'BOM记录数': len(self.df),
            '物料总数': len(set(self.df['父物料编码'].unique()) | set(self.df['子物料编码'].unique())),
            '最大层级': int(self.df['层级'].max()),
            '最小层级': int(self.df['层级'].min()),
            '平均生产周期': round(self.df['生产周期(天)'].mean(), 2)
        }
    
    def get_materials_by_level(self, level):
        """
        获取指定层级的物料
        
        Args:
            level: 层级编号
            
        Returns:
            list: 物料编码列表
        """
        if self.df is None:
            return []
        
        return self.df[self.df['层级'] == level]['子物料编码'].unique().tolist()
    
    def get_children(self, material_code):
        """
        获取指定物料的所有子物料
        
        Args:
            material_code: 物料编码
            
        Returns:
            pandas DataFrame: 子物料信息
        """
        if self.df is None:
            return pd.DataFrame()
        
        return self.df[self.df['父物料编码'] == material_code]
