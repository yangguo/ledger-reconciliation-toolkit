import pandas as pd
import numpy as np
from datetime import datetime

def parse_currency_value(value):
    """解析货币值，处理逗号分隔符和负号"""
    # 处理 pandas Series 或单个值
    try:
        # 如果是 Series，取第一个值
        if hasattr(value, 'iloc'):
            value = value.iloc[0] if len(value) > 0 else None
        
        if value is None or pd.isna(value) or value == '':
            return 0.0
    except:
        # 如果出现任何错误，返回 0.0
        return 0.0
    
    # 转换为字符串
    str_value = str(value).strip()
    
    # 处理标题行和非数字内容
    if str_value in ['本币', '原币', '币种', '科目编码', 'nan', 'NaN']:
        return 0.0
    
    # 移除逗号
    str_value = str_value.replace(',', '')
    
    # 处理负号格式 "- 123.45" -> "-123.45"
    if str_value.startswith('- '):
        str_value = '-' + str_value[2:]
    
    try:
        return float(str_value)
    except ValueError:
        print(f"警告: 无法解析货币值 '{value}'，返回0")
        return 0.0

def extract_account_code(subject_field):
    """从科目字段中提取科目编码"""
    if pd.isna(subject_field):
        return ''
    
    subject_str = str(subject_field).strip()
    
    # 如果包含反斜杠，说明是JE格式，提取第一部分作为科目编码
    if '\\' in subject_str:
        parts = subject_str.split('\\')
        return parts[0].strip()
    
    # 对于TB数据，直接返回原值（不进行数字提取，保留完整科目编码）
    # 这样可以保留像11330102A8这样包含字母的科目编码
    return subject_str