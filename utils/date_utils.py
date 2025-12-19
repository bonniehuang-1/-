# -*- coding: utf-8 -*-
"""
日期工具类
"""
import pandas as pd
from datetime import datetime, timedelta
from config import Config


class DateUtils:
    """日期处理工具类"""
    
    @staticmethod
    def is_workday(date):
        """
        判断是否为工作日
        
        Args:
            date: datetime对象或pandas.Timestamp
            
        Returns:
            bool: True表示工作日，False表示非工作日
        """
        # 转换为datetime对象
        if isinstance(date, pd.Timestamp):
            date = date.to_pydatetime()
        
        # 检查是否为周末
        if date.weekday() not in Config.WORK_DAYS:
            return False
        
        # 检查是否为节假日
        date_str = date.strftime('%Y-%m-%d')
        if date_str in Config.HOLIDAYS:
            return False
        
        return True
    
    @staticmethod
    def add_workdays(start_date, days):
        """
        从起始日期开始，增加指定的工作日数
        
        Args:
            start_date: 起始日期
            days: 要增加的工作日数
            
        Returns:
            datetime: 计算后的日期
        """
        if isinstance(start_date, pd.Timestamp):
            start_date = start_date.to_pydatetime()
        
        current_date = start_date
        workdays_added = 0
        
        while workdays_added < days:
            current_date += timedelta(days=1)
            if DateUtils.is_workday(current_date):
                workdays_added += 1
        
        return current_date
    
    @staticmethod
    def subtract_workdays(end_date, days):
        """
        从结束日期开始，减去指定的工作日数
        
        Args:
            end_date: 结束日期
            days: 要减去的工作日数
            
        Returns:
            datetime: 计算后的日期
        """
        if isinstance(end_date, pd.Timestamp):
            end_date = end_date.to_pydatetime()
        
        current_date = end_date
        workdays_subtracted = 0
        
        while workdays_subtracted < days:
            current_date -= timedelta(days=1)
            if DateUtils.is_workday(current_date):
                workdays_subtracted += 1
        
        return current_date
    
    @staticmethod
    def count_workdays(start_date, end_date):
        """
        计算两个日期之间的工作日数
        
        Args:
            start_date: 起始日期
            end_date: 结束日期
            
        Returns:
            int: 工作日数
        """
        if isinstance(start_date, pd.Timestamp):
            start_date = start_date.to_pydatetime()
        if isinstance(end_date, pd.Timestamp):
            end_date = end_date.to_pydatetime()
        
        if start_date > end_date:
            return 0
        
        workdays = 0
        current_date = start_date
        
        while current_date <= end_date:
            if DateUtils.is_workday(current_date):
                workdays += 1
            current_date += timedelta(days=1)
        
        return workdays
    
    @staticmethod
    def get_next_workday(date):
        """
        获取下一个工作日
        
        Args:
            date: 当前日期
            
        Returns:
            datetime: 下一个工作日
        """
        if isinstance(date, pd.Timestamp):
            date = date.to_pydatetime()
        
        next_day = date + timedelta(days=1)
        while not DateUtils.is_workday(next_day):
            next_day += timedelta(days=1)
        
        return next_day
    
    @staticmethod
    def format_date(date):
        """
        格式化日期为字符串
        
        Args:
            date: datetime对象或pandas.Timestamp
            
        Returns:
            str: 格式化的日期字符串 (YYYY-MM-DD)
        """
        if isinstance(date, pd.Timestamp):
            date = date.to_pydatetime()
        return date.strftime('%Y-%m-%d')
