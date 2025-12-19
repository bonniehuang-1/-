# -*- coding: utf-8 -*-
"""
数据验证工具类
"""
import pandas as pd
from config import Config


class DataValidator:
    """数据验证工具类"""
    
    @staticmethod
    def validate_columns(df, required_columns, data_type="数据"):
        """
        验证DataFrame是否包含所有必填列
        
        Args:
            df: pandas DataFrame
            required_columns: 必填列名列表
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果缺少必填列
        """
        missing_columns = set(required_columns) - set(df.columns)
        if missing_columns:
            raise ValueError(f"{data_type}缺少必填列: {', '.join(missing_columns)}")
    
    @staticmethod
    def validate_date_logic(df, start_col, end_col, data_type="数据"):
        """
        验证日期逻辑（开始日期应早于结束日期）
        
        Args:
            df: pandas DataFrame
            start_col: 开始日期列名
            end_col: 结束日期列名
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果存在日期逻辑错误
        """
        invalid_dates = df[df[start_col] > df[end_col]]
        if not invalid_dates.empty:
            error_records = invalid_dates.index.tolist()
            raise ValueError(
                f"{data_type}中发现{len(invalid_dates)}条记录的{start_col}晚于{end_col}，"
                f"记录索引: {error_records[:5]}{'...' if len(error_records) > 5 else ''}"
            )
    
    @staticmethod
    def validate_positive_numbers(df, columns, data_type="数据"):
        """
        验证数值列是否为正数
        
        Args:
            df: pandas DataFrame
            columns: 需要验证的列名列表
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果存在非正数
        """
        for col in columns:
            if col in df.columns:
                invalid_values = df[df[col] <= Config.MIN_QUANTITY]
                if not invalid_values.empty:
                    error_records = invalid_values.index.tolist()
                    raise ValueError(
                        f"{data_type}中列'{col}'存在{len(invalid_values)}条非正数记录，"
                        f"记录索引: {error_records[:5]}{'...' if len(error_records) > 5 else ''}"
                    )
    
    @staticmethod
    def validate_no_nulls(df, columns, data_type="数据"):
        """
        验证指定列是否存在空值
        
        Args:
            df: pandas DataFrame
            columns: 需要验证的列名列表
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果存在空值
        """
        for col in columns:
            if col in df.columns:
                null_count = df[col].isnull().sum()
                if null_count > 0:
                    raise ValueError(
                        f"{data_type}中列'{col}'存在{null_count}个空值"
                    )
    
    @staticmethod
    def validate_unique(df, columns, data_type="数据"):
        """
        验证指定列的值是否唯一
        
        Args:
            df: pandas DataFrame
            columns: 需要验证的列名列表
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果存在重复值
        """
        for col in columns:
            if col in df.columns:
                duplicates = df[df.duplicated(subset=[col], keep=False)]
                if not duplicates.empty:
                    dup_values = duplicates[col].unique().tolist()
                    raise ValueError(
                        f"{data_type}中列'{col}'存在{len(dup_values)}个重复值: "
                        f"{dup_values[:5]}{'...' if len(dup_values) > 5 else ''}"
                    )
    
    @staticmethod
    def validate_data_types(df, type_mapping, data_type="数据"):
        """
        验证并转换数据类型
        
        Args:
            df: pandas DataFrame
            type_mapping: 字典，键为列名，值为目标数据类型
            data_type: 数据类型名称（用于错误提示）
            
        Returns:
            pandas DataFrame: 转换后的DataFrame
            
        Raises:
            ValueError: 如果数据类型转换失败
        """
        for col, dtype in type_mapping.items():
            if col in df.columns:
                try:
                    if dtype == 'datetime':
                        df[col] = pd.to_datetime(df[col])
                    elif dtype == 'int':
                        df[col] = df[col].astype(int)
                    elif dtype == 'float':
                        df[col] = df[col].astype(float)
                    elif dtype == 'str':
                        df[col] = df[col].astype(str)
                except Exception as e:
                    raise ValueError(
                        f"{data_type}中列'{col}'转换为{dtype}类型失败: {str(e)}"
                    )
        return df
    
    @staticmethod
    def validate_reference_integrity(df, ref_col, ref_df, ref_key_col, data_type="数据"):
        """
        验证引用完整性（外键约束）
        
        Args:
            df: 主DataFrame
            ref_col: 主DataFrame中的引用列
            ref_df: 参考DataFrame
            ref_key_col: 参考DataFrame中的键列
            data_type: 数据类型名称（用于错误提示）
            
        Raises:
            ValueError: 如果存在无效引用
        """
        if ref_col in df.columns and ref_key_col in ref_df.columns:
            valid_refs = set(ref_df[ref_key_col].unique())
            invalid_refs = df[~df[ref_col].isin(valid_refs)]
            
            if not invalid_refs.empty:
                invalid_values = invalid_refs[ref_col].unique().tolist()
                raise ValueError(
                    f"{data_type}中列'{ref_col}'存在{len(invalid_values)}个无效引用: "
                    f"{invalid_values[:5]}{'...' if len(invalid_values) > 5 else ''}"
                )
