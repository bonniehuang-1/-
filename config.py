# -*- coding: utf-8 -*-
"""
生产排产系统 - 配置文件
"""
import os

class Config:
    """系统配置类"""
    
    # 项目根目录
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    
    # 文件路径配置
    INPUT_DIR = os.path.join(BASE_DIR, "input")
    OUTPUT_DIR = os.path.join(BASE_DIR, "output")
    TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")
    LOG_DIR = os.path.join(BASE_DIR, "logs")
    
    # 输入文件名
    ORDERS_FILE = "orders.xlsx"
    BOM_FILE = "bom.xlsx"
    CAPACITY_FILE = "capacity.xlsx"
    
    # 输出文件名
    REPORT_FILE = "delivery_analysis_report.xlsx"
    ALERT_FILE = "delivery_alerts.xlsx"
    
    # 业务规则配置
    WORK_DAYS = [0, 1, 2, 3, 4]  # 周一到周五 (0=周一, 6=周日)
    HOLIDAYS = []  # 节假日列表，格式：['2024-01-01', '2024-02-10']
    
    # 预警阈值
    ALERT_DELAY_DAYS_RED = 7    # 红色预警：延期>=7天
    ALERT_DELAY_DAYS_YELLOW = 1  # 黄色预警：延期1-6天
    CAPACITY_UTILIZATION_THRESHOLD = 0.9  # 产能利用率预警阈值（90%）
    
    # 日志配置
    LOG_LEVEL = "INFO"  # DEBUG, INFO, WARNING, ERROR, CRITICAL
    LOG_FILE = "production_scheduler.log"
    LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    
    # 数据验证配置
    MAX_BOM_LEVEL = 10  # 最大BOM层级深度
    MIN_QUANTITY = 0.001  # 最小数量（避免浮点数精度问题）
    
    @classmethod
    def ensure_directories(cls):
        """确保所有必要的目录存在"""
        for directory in [cls.INPUT_DIR, cls.OUTPUT_DIR, cls.TEMPLATE_DIR, cls.LOG_DIR]:
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"创建目录: {directory}")
    
    @classmethod
    def get_input_file_path(cls, filename):
        """获取输入文件的完整路径"""
        return os.path.join(cls.INPUT_DIR, filename)
    
    @classmethod
    def get_output_file_path(cls, filename):
        """获取输出文件的完整路径"""
        return os.path.join(cls.OUTPUT_DIR, filename)
    
    @classmethod
    def get_log_file_path(cls):
        """获取日志文件的完整路径"""
        return os.path.join(cls.LOG_DIR, cls.LOG_FILE)
