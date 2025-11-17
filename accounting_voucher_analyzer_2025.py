#!/usr/bin/env python3
"""
会计分录检查脚本 - Accounting Voucher Analysis Tool
对会计分录文件进行全面的合规性检查，支持2022、2023、2024、2025年数据
使用openpyxl和pandas进行xlsx文件分析
"""

import pandas as pd
import openpyxl
import re
import sys
import os
from datetime import datetime
import calendar
from collections import Counter, defaultdict

class AccountingVoucherAnalyzer:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.vouchers = []
        self.working_days = set()
        
        # 2022年法定节假日（放假日期）
        self.chinese_holidays_2022 = {
            '2022-01-01', '2022-01-02', '2022-01-03',  # 元旦
            '2022-01-31', '2022-02-01', '2022-02-02', '2022-02-03', '2022-02-04', '2022-02-05', '2022-02-06',  # 春节
            '2022-04-03', '2022-04-04', '2022-04-05',  # 清明节
            '2022-04-30', '2022-05-01', '2022-05-02', '2022-05-03', '2022-05-04',  # 劳动节
            '2022-06-03', '2022-06-04', '2022-06-05',  # 端午节
            '2022-09-10', '2022-09-11', '2022-09-12',  # 中秋节
            '2022-10-01', '2022-10-02', '2022-10-03', '2022-10-04', '2022-10-05', '2022-10-06', '2022-10-07'  # 国庆节
        }
        
        # 2022年调休工作日
        self.chinese_makeup_workdays_2022 = {
            '2022-01-29', '2022-01-30',  # 春节前调休
            '2022-04-02', '2022-04-24',  # 清明节和劳动节调休
            '2022-10-08', '2022-10-09'   # 国庆节后调休
        }
        
        # 2023年法定节假日（放假日期）
        self.chinese_holidays_2023 = {
            '2023-01-01', '2023-01-02',  # 元旦
            '2023-01-21', '2023-01-22', '2023-01-23', '2023-01-24', '2023-01-25', '2023-01-26', '2023-01-27',  # 春节
            '2023-04-05',  # 清明节
            '2023-05-01', '2023-05-02', '2023-05-03',  # 劳动节
            '2023-06-22', '2023-06-23', '2023-06-24',  # 端午节
            '2023-09-29', '2023-09-30', '2023-10-01', '2023-10-02', '2023-10-03', '2023-10-04', '2023-10-05', '2023-10-06'  # 中秋国庆
        }
        
        # 2023年调休工作日
        self.chinese_makeup_workdays_2023 = {
            '2023-01-28', '2023-01-29',  # 春节后调休
            '2023-04-23',  # 劳动节前调休
            '2023-05-06',  # 劳动节后调休
            '2023-06-25',  # 端午节后调休
            '2023-10-07', '2023-10-08'   # 国庆节后调休
        }
        
        # 2024年法定节假日（放假日期）
        self.chinese_holidays_2024 = {
            '2024-01-01',  # 元旦
            '2024-02-10', '2024-02-11', '2024-02-12', '2024-02-13', '2024-02-14', '2024-02-15', '2024-02-16', '2024-02-17',  # 春节
            '2024-04-04', '2024-04-05', '2024-04-06',  # 清明节
            '2024-05-01', '2024-05-02', '2024-05-03', '2024-05-04', '2024-05-05',  # 劳动节
            '2024-06-08', '2024-06-09', '2024-06-10',  # 端午节
            '2024-09-15', '2024-09-16', '2024-09-17',  # 中秋节
            '2024-10-01', '2024-10-02', '2024-10-03', '2024-10-04', '2024-10-05', '2024-10-06', '2024-10-07'  # 国庆节
        }
        
        # 2024年调休工作日
        self.chinese_makeup_workdays_2024 = {
            '2024-02-04', '2024-02-18',  # 春节调休
            '2024-04-07', '2024-04-28',  # 清明节和劳动节调休
            '2024-05-11',  # 劳动节后调休
            '2024-09-14', '2024-09-29',  # 中秋节和国庆节调休
            '2024-10-12'   # 国庆节后调休
        }
        
        # 2025年法定节假日（放假日期）
        self.chinese_holidays_2025 = {
            '2025-01-01',  # 元旦
            '2025-01-29', '2025-01-30', '2025-01-31', '2025-02-01', '2025-02-02', '2025-02-03', '2025-02-04',  # 春节
            '2025-04-04', '2025-04-05', '2025-04-06',  # 清明节
            '2025-05-01', '2025-05-02', '2025-05-03', '2025-05-04',  # 劳动节
            '2025-05-31', '2025-06-01', '2025-06-02',  # 端午节
            '2025-10-01', '2025-10-02', '2025-10-03', '2025-10-04', '2025-10-05', '2025-10-06', '2025-10-07',  # 国庆节
            '2025-10-08'  # 中秋节
        }
        
        # 2025年调休工作日
        self.chinese_makeup_workdays_2025 = {
            '2025-01-26', '2025-02-08',  # 春节调休
            '2025-04-27', '2025-09-28',  # 劳动节和国庆节调休
            '2025-10-11'  # 国庆节后调休
        }
        
        # 合并所有年份的节假日和调休工作日
        self.all_holidays = set()
        self.all_holidays.update(self.chinese_holidays_2022)
        self.all_holidays.update(self.chinese_holidays_2023)
        self.all_holidays.update(self.chinese_holidays_2024)
        self.all_holidays.update(self.chinese_holidays_2025)
        
        self.all_makeup_workdays = set()
        self.all_makeup_workdays.update(self.chinese_makeup_workdays_2022)
        self.all_makeup_workdays.update(self.chinese_makeup_workdays_2023)
        self.all_makeup_workdays.update(self.chinese_makeup_workdays_2024)
        self.all_makeup_workdays.update(self.chinese_makeup_workdays_2025)
        
    def safe_get_field(self, record, field_name, default=''):
        """安全获取记录字段，避免KeyError"""
        try:
            if isinstance(record, dict):
                return record.get(field_name, default)
            else:
                return default
        except Exception as e:
            print(f"⚠️ 获取字段 {field_name} 时出错: {e}")
            return default
            
    def validate_record_structure(self, record):
        """验证记录结构的完整性"""
        required_fields = ['凭证号', '日期', '制单人', '审核人']
        missing_fields = []
        
        for field in required_fields:
            if field not in record or record[field] is None:
                missing_fields.append(field)
                
        return missing_fields
        
    def safe_print_record(self, record, format_string, *field_names):
        """安全打印记录信息"""
        try:
            values = []
            for field_name in field_names:
                value = self.safe_get_field(record, field_name, '')
                values.append(value)
            print(format_string.format(*values))
        except Exception as e:
            print(f"⚠️ 打印记录时出错: {e}")
            print(f"   记录内容: {record}")
        
    def set_file_path(self, file_path):
        """设置要分析的文件路径"""
        self.file_path = file_path
        self.vouchers = []  # 重置凭证数据
        
    def is_weekday(self, date_str):
        """检查是否为工作日（考虑调休安排）"""
        try:
            date_obj = datetime.strptime(date_str, '%Y-%m-%d')
            date_str_fmt = date_obj.strftime('%Y-%m-%d')
            
            # 检查是否为调休工作日（原本是周末但需要上班）
            if date_str_fmt in self.all_makeup_workdays:
                return True
                
            # 检查是否为法定节假日
            if date_str_fmt in self.all_holidays:
                return False
            
            # 检查是否为周末
            if date_obj.weekday() >= 5:  # 周六=5, 周日=6
                return False
                
            return True
        except:
            return None
    
    def parse_excel_data(self):
        """使用pandas和openpyxl解析Excel文件数据，优化内存使用"""
        try:
            # 使用pandas读取Excel文件
            workbook = pd.ExcelFile(self.file_path)
            
            # 获取所有工作表名称
            sheet_names = workbook.sheet_names
            print(f"📊 发现 {len(sheet_names)} 个工作表: {sheet_names}")
            
            for sheet_name in sheet_names:
                print(f"🔍 正在处理工作表: {sheet_name}")
                
                # 使用更小的chunk_size和优化的读取方式
                chunk_size = 50000  # 减少到1000行
                
                try:
                    # 先读取一小部分来获取列信息和总行数
                    print(f"📊 正在获取工作表信息...")
                    df_sample = pd.read_excel(self.file_path, sheet_name=sheet_name, nrows=1)
                    
                    # 使用openpyxl直接获取行数，避免读取整个文件
                    import openpyxl
                    wb = openpyxl.load_workbook(self.file_path, read_only=True)
                    ws = wb[sheet_name]
                    total_rows = ws.max_row - 1  # 减去标题行
                    wb.close()
                    
                    print(f"📈 工作表 {sheet_name} 总行数: {total_rows}")
                    
                    # 分块处理数据
                    processed_rows = 0
                    for i in range(0, total_rows, chunk_size):
                        current_chunk_size = min(chunk_size, total_rows - i)
                        if (i // chunk_size) % 10 == 0:  # 每10个chunk打印一次进度
                            print(f"📖 读取第 {i+1}-{i+current_chunk_size} 行 ({processed_rows/total_rows*100:.1f}%)")
                        
                        try:
                            df_chunk = pd.read_excel(
                                self.file_path, 
                                sheet_name=sheet_name,
                                skiprows=range(1, i+1) if i > 0 else None,
                                nrows=current_chunk_size,
                                engine='openpyxl'
                            )
                            
                            if df_chunk.empty:
                                continue
                            
                            # 标准化列名（去除空格和特殊字符）
                            df_chunk.columns = df_chunk.columns.astype(str).str.strip()
                            
                            # 映射常见列名到标准字段
                            column_mapping = {
                                '凭证号': '凭证号',
                                '凭证编号': '凭证号',
                                '日期': '日期',
                                '凭证日期': '日期',
                                '摘要': '摘要',
                                '凭证摘要': '摘要',
                                '制单人': '制单人',
                                '制单': '制单人',
                                '录入人': '制单人',
                                '审核人': '审核人',
                                '审核': '审核人',
                                '过账人': '过账人',
                                '过账': '过账人',
                                '记账人': '过账人',
                                '金额': '金额',
                                '借方金额': '借方金额',
                                '贷方金额': '贷方金额',
                                '借方原币': '借方原币',
                                '借方本币': '借方本币',
                                '贷方原币': '贷方原币',
                                '贷方本币': '贷方本币',
                                '科目': '科目',
                                '科目名称': '科目',
                                '会计科目': '科目',
                                '年': '年',
                                '月': '月',
                                '账簿': '账簿',
                                '分录号': '分录号',
                                '币种': '币种',
                                '来源系统': '来源系统'
                            }
                            
                            # 重命名列
                            df_renamed = df_chunk.rename(columns=column_mapping)
                            
                            # 将数据转换为字典列表
                            for idx, row in df_renamed.iterrows():
                                voucher_data = {
                                    'sheet': sheet_name,
                                    'row': int(idx) + 2 + i,  # 计算实际行号
                                }
                                
                                # 包含用户要求的完整16个JE原始字段
                                required_fields = [
                                    '年', '月', '账簿', '凭证号', '分录号', '摘要', '科目', '币种',
                                    '借方原币', '借方本币', '贷方原币', '贷方本币', '来源系统',
                                    '制单人', '审核人', '过账人'
                                ]
                                
                                # 添加所有可用字段，优先保存必需字段
                                for col in df_renamed.columns:
                                    voucher_data[col] = str(row[col]).strip() if pd.notna(row[col]) else ''
                                
                                # 确保所有必需字段存在
                                for field in required_fields:
                                    if field not in voucher_data:
                                        voucher_data[field] = ''
                                
                                # 为了兼容性，保留原有的金额字段
                                if '金额' not in voucher_data and ('借方本币' in voucher_data or '贷方本币' in voucher_data):
                                    debit = float(voucher_data.get('借方本币', 0) or 0)
                                    credit = float(voucher_data.get('贷方本币', 0) or 0)
                                    voucher_data['金额'] = str(max(debit, credit))
                                
                                self.vouchers.append(voucher_data)
                            
                            # 清理内存并更新进度
                            del df_chunk
                            del df_renamed
                            processed_rows += current_chunk_size
                            
                            # 强制垃圾回收以释放内存
                            if processed_rows % (chunk_size * 5) == 0:
                                import gc
                                gc.collect()
                             
                        except Exception as chunk_error:
                            # 减少错误输出，只记录关键错误
                            if "Memory" in str(chunk_error) or "pandas" in str(chunk_error):
                                print(f"⚠️ 处理数据块时出错: {str(chunk_error)}")
                            continue
                    
                    # 减少输出，只在处理大量数据时显示
                    if total_rows > 1000:
                        print(f"✅ 工作表 {sheet_name} 处理完成，共 {len([v for v in self.vouchers if v['sheet'] == sheet_name])} 条记录")
                    
                except Exception as sheet_error:
                    # 只输出关键错误信息
                    if "permission" in str(sheet_error).lower() or "file" in str(sheet_error).lower():
                        print(f"⚠️ 处理工作表 {sheet_name} 时出错: {str(sheet_error)}")
                    continue
                        
        except Exception as e:
            print(f"❌ Error parsing Excel: {str(e)}")
            return False
        
        print(f"🎉 Excel解析完成，总共处理 {len(self.vouchers)} 条记录")
        return True
    
    def check_duplicate_maker_reviewer(self):
        """1. 检查制单人和审核人为同一人的情况"""
        results = []
        for voucher in self.vouchers:
            maker = self.safe_get_field(voucher, '制单人')
            reviewer = self.safe_get_field(voucher, '审核人')
            if maker and reviewer and maker == reviewer:
                # 创建包含完整JE原始字段的记录
                result = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '制单审核同一人'
                }
                results.append(result)
        return results
    
    def check_unauthorized_makers(self):
        """2. 检查IT人员、刘盛艳、罗贻芬制单的分录"""
        unauthorized_names = ['刘盛艳', '罗贻芬', '罗貽芬']
        # IT人员常见名称模式
        it_patterns = [r'IT', r'技术', r'信息', r'系统', r'管理员']
        
        results = []
        for voucher in self.vouchers:
            maker = self.safe_get_field(voucher, '制单人')
            if not maker:
                continue
                
            violation_type = None
            # 检查具体人员
            if any(name in maker for name in unauthorized_names):
                violation_type = '高管制单'
            else:
                # 检查IT人员
                for pattern in it_patterns:
                    if re.search(pattern, maker, re.IGNORECASE):
                        violation_type = 'IT人员制单'
                        break
            
            if violation_type:
                # 创建包含完整JE原始字段的记录
                result = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': violation_type
                }
                results.append(result)
        
        return results
    
    def check_empty_descriptions(self):
        """3. 检查没有摘要的分录"""
        results = []
        for voucher in self.vouchers:
            description = self.safe_get_field(voucher, '摘要')
            if not description or description.strip() == '':
                # 创建包含完整JE原始字段的记录
                result = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', '') or '空',
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '无摘要分录'
                }
                results.append(result)
        return results
    
    def check_personnel_completeness(self):
        """4. 检查编制人为空的情况并统计所有人员"""
        personnel = {
            '制单人': set(),
            '审核人': set(),
            '过账人': set()
        }
        
        empty_fields = []
        
        for voucher in self.vouchers:
            # 统计所有人员
            maker = self.safe_get_field(voucher, '制单人')
            reviewer = self.safe_get_field(voucher, '审核人')
            poster = self.safe_get_field(voucher, '过账人')
            if maker:
                personnel['制单人'].add(maker)
            if reviewer:
                personnel['审核人'].add(reviewer)
            if poster:
                personnel['过账人'].add(poster)
            
            # 检查空值
            maker = self.safe_get_field(voucher, '制单人')
            if not maker or maker.strip() == '':
                empty_record = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '制单人信息空值',
                    '空值字段': '制单人'
                }
                empty_fields.append(empty_record)
            reviewer = self.safe_get_field(voucher, '审核人')
            if not reviewer or reviewer.strip() == '':
                empty_record = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '审核人信息空值',
                    '空值字段': '审核人'
                }
                empty_fields.append(empty_record)
            poster = self.safe_get_field(voucher, '过账人')
            if not poster or poster.strip() == '':
                empty_record = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '过账人信息空值',
                    '空值字段': '过账人'
                }
                empty_fields.append(empty_record)
        
        return {
            '所有人员': {k: list(v) for k, v in personnel.items()},
            '空值记录': empty_fields
        }
    
    def check_adjustment_vouchers(self):
        """5. 检查摘要中包含调整的分录"""
        # 精简后的关键词列表，只保留真正需要关注的异常调整关键词
        adjustment_keywords = [
            '调整', '更正', '纠正', '修正', '修改', '冲正', '调账', 
            '重分类', '冲销', '冲回', '差异', '差额'
        ]
        
        results = []
        for voucher in self.vouchers:
            description = self.safe_get_field(voucher, '摘要')
            matched_keywords = []
            
            # 检查所有匹配的关键词
            for keyword in adjustment_keywords:
                if keyword in description:
                    matched_keywords.append(keyword)
            
            if matched_keywords:
                # 创建包含完整JE原始字段的调整分录记录
                adjustment_record = {
                    '年': voucher.get('年', ''),
                    '月': voucher.get('月', ''),
                    '账簿': voucher.get('账簿', ''),
                    '凭证号': voucher.get('凭证号', ''),
                    '分录号': voucher.get('分录号', ''),
                    '摘要': voucher.get('摘要', ''),
                    '科目': voucher.get('科目', ''),
                    '币种': voucher.get('币种', ''),
                    '借方原币': voucher.get('借方原币', ''),
                    '借方本币': voucher.get('借方本币', ''),
                    '贷方原币': voucher.get('贷方原币', ''),
                    '贷方本币': voucher.get('贷方本币', ''),
                    '来源系统': voucher.get('来源系统', ''),
                    '制单人': voucher.get('制单人', ''),
                    '审核人': voucher.get('审核人', ''),
                    '过账人': voucher.get('过账人', ''),
                    '违规类型': '调整类分录',
                    '关键词': ', '.join(matched_keywords)
                }
                
                results.append(adjustment_record)
        
        return results
    
    def extract_name_from_rpa(self, rpa_name):
        """
        从RPA制单人名称中提取真实姓名
        例如："邓鹏程RPA" -> "邓鹏程"
        """
        # 移除RPA、RPA2等后缀
        name = re.sub(r'RPA\d*$', '', rpa_name)
        # 移除可能的其他后缀
        name = name.strip()
        return name
    
    def is_rpa_maker(self, maker_name):
        """
        判断是否为RPA制单人
        """
        if not maker_name:
            return False
        return 'RPA' in str(maker_name) or '自动化' in str(maker_name)
    
    def check_rpa_reviewer_compliance(self):
        """
        7. 检查RPA制单人与审核人的合规性
        """
        # 统计所有RPA制单人
        rpa_makers = set()
        rpa_combinations = []
        non_compliant_cases = []
        
        for voucher in self.vouchers:
            maker = self.safe_get_field(voucher, '制单人')
            reviewer = self.safe_get_field(voucher, '审核人')
            
            if not maker or not reviewer:
                continue
                
            if self.is_rpa_maker(maker):
                rpa_makers.add(maker)
                rpa_combinations.append((maker, reviewer))
                
                # 提取RPA制单人中的真实姓名
                real_name = self.extract_name_from_rpa(maker)
                
                # 检查审核人是否与真实姓名相同
                if real_name == reviewer:
                    # 创建包含完整JE原始字段的RPA不合规记录
                    non_compliant_record = {
                        '年': voucher.get('年', ''),
                        '月': voucher.get('月', ''),
                        '账簿': voucher.get('账簿', ''),
                        '凭证号': voucher.get('凭证号', ''),
                        '分录号': voucher.get('分录号', ''),
                        '摘要': voucher.get('摘要', ''),
                        '科目': voucher.get('科目', ''),
                        '币种': voucher.get('币种', ''),
                        '借方原币': voucher.get('借方原币', ''),
                        '借方本币': voucher.get('借方本币', ''),
                        '贷方原币': voucher.get('贷方原币', ''),
                        '贷方本币': voucher.get('贷方本币', ''),
                        '来源系统': voucher.get('来源系统', ''),
                        '制单人': voucher.get('制单人', ''),
                        '审核人': voucher.get('审核人', ''),
                        '过账人': voucher.get('过账人', ''),
                        '违规类型': 'RPA合规性检查',
                        '提取的真实姓名': real_name,
                        '风险等级': '高风险',
                        '合规状态': '不合规',
                        '问题描述': 'RPA制单人名称中的姓名与审核人相同'
                    }
                    
                    non_compliant_cases.append(non_compliant_record)
        
        # 创建RPA制单人对应审核人的分析数据
        rpa_reviewer_mapping = defaultdict(set)
        for maker, reviewer in rpa_combinations:
            rpa_reviewer_mapping[maker].add(reviewer)
        
        rpa_analysis_data = []
        for rpa_maker, reviewers in rpa_reviewer_mapping.items():
            real_name = self.extract_name_from_rpa(rpa_maker)
            reviewers_list = sorted(list(reviewers))
            has_same_name = real_name in reviewers_list
            
            rpa_analysis_data.append({
                'RPA制单人': rpa_maker,
                '提取的真实姓名': real_name,
                '审核人列表': ', '.join(reviewers_list),
                '审核人数量': len(reviewers_list),
                '包含同名审核人': '是' if has_same_name else '否',
                '风险状态': '高风险' if has_same_name else '正常'
            })
        
        return {
            'rpa_makers': list(rpa_makers),
            'rpa_combinations_count': len(rpa_combinations),
            'non_compliant_cases': non_compliant_cases,
            'rpa_analysis_data': rpa_analysis_data
        }
    
    def analyze_maker_reviewer_combinations(self):
        """
        8. 分析制单人审核人组合关系
        """
        # 创建组合统计
        combination_counter = Counter()
        maker_reviewers = defaultdict(set)
        reviewer_makers = defaultdict(set)
        same_person_count = 0
        
        for voucher in self.vouchers:
            maker = self.safe_get_field(voucher, '制单人')
            reviewer = self.safe_get_field(voucher, '审核人')
            
            if not maker or not reviewer or maker == 'nan' or reviewer == 'nan':
                continue
                
            combination_key = f"{maker} → {reviewer}"
            combination_counter[combination_key] += 1
            maker_reviewers[maker].add(reviewer)
            reviewer_makers[reviewer].add(maker)
            
            # 统计制单审核同一人
            if maker == reviewer:
                same_person_count += 1
        
        # 创建详细组合数据
        combination_data = []
        total_combinations = sum(combination_counter.values())
        
        for combo_key, count in combination_counter.items():
            maker, reviewer = combo_key.split(' → ')
            combination_data.append({
                '制单人': maker,
                '审核人': reviewer,
                '组合次数': count,
                '占比(%)': round((count / total_combinations) * 100, 2),
                '是否同一人': '是' if maker == reviewer else '否'
            })
        
        # 制单人统计数据
        maker_reviewer_counts = [(maker, len(reviewers)) for maker, reviewers in maker_reviewers.items()]
        maker_reviewer_counts.sort(key=lambda x: x[1], reverse=True)
        
        maker_data = []
        for maker, count in maker_reviewer_counts:
            reviewers_list = ', '.join(sorted(maker_reviewers[maker]))
            maker_data.append({
                '制单人': maker,
                '审核人数量': count,
                '审核人列表': reviewers_list
            })
        
        # 审核人统计数据
        reviewer_maker_counts = [(reviewer, len(makers)) for reviewer, makers in reviewer_makers.items()]
        reviewer_maker_counts.sort(key=lambda x: x[1], reverse=True)
        
        reviewer_data = []
        for reviewer, count in reviewer_maker_counts:
            makers_list = ', '.join(sorted(reviewer_makers[reviewer]))
            reviewer_data.append({
                '审核人': reviewer,
                '制单人数量': count,
                '制单人列表': makers_list
            })
        
        # 制单审核同一人详细数据
        same_person_data = []
        if same_person_count > 0:
            same_person_counter = Counter()
            for voucher in self.vouchers:
                maker = self.safe_get_field(voucher, '制单人')
                reviewer = self.safe_get_field(voucher, '审核人')
                if maker and reviewer and maker == reviewer:
                    same_person_counter[maker] += 1
            
            for person, count in same_person_counter.items():
                same_person_data.append({
                    '人员': person,
                    '同一人次数': count,
                    '占同一人总数比例(%)': round((count / same_person_count) * 100, 2)
                })
        
        return {
            'total_combinations': total_combinations,
            'unique_combinations': len(combination_counter),
            'same_person_count': same_person_count,
            'unique_makers': len(maker_reviewers),
            'unique_reviewers': len(reviewer_makers),
            'combination_data': sorted(combination_data, key=lambda x: x['组合次数'], reverse=True),
            'maker_data': maker_data,
            'reviewer_data': reviewer_data,
            'same_person_data': sorted(same_person_data, key=lambda x: x['同一人次数'], reverse=True),
            'top_combinations': combination_counter.most_common(20)
        }
    
    def check_weekend_vouchers(self):
        """6. 检查非工作日制单的凭证（包括调休工作日分析）"""
        results = []
        makeup_workday_results = []
        
        for voucher in self.vouchers:
            date_str = voucher.get('日期', '')
            if not date_str or str(date_str).strip() == '':
                continue
            
            # 尝试解析日期
            try:
                # 处理不同日期格式
                date_obj = None
                for fmt in ['%Y-%m-%d', '%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y']:
                    try:
                        date_obj = datetime.strptime(date_str.strip(), fmt)
                        break
                    except ValueError:
                        continue
                
                if date_obj is None:
                    continue
                
                date_str_fmt = date_obj.strftime('%Y-%m-%d')
                day_name = calendar.day_name[date_obj.weekday()]
                
                # 检查是否为调休工作日（原本是周末但需要上班）
                if date_str_fmt in self.all_makeup_workdays:
                    makeup_workday_results.append({
                        '年': voucher.get('年', ''),
                        '月': voucher.get('月', ''),
                        '账簿': voucher.get('账簿', ''),
                        '凭证号': voucher.get('凭证号', ''),
                        '分录号': voucher.get('分录号', ''),
                        '摘要': voucher.get('摘要', ''),
                        '科目': voucher.get('科目', ''),
                        '币种': voucher.get('币种', ''),
                        '借方原币': voucher.get('借方原币', ''),
                        '借方本币': voucher.get('借方本币', ''),
                        '贷方原币': voucher.get('贷方原币', ''),
                        '贷方本币': voucher.get('贷方本币', ''),
                        '来源系统': voucher.get('来源系统', ''),
                        '制单人': voucher.get('制单人', ''),
                        '审核人': voucher.get('审核人', ''),
                        '过账人': voucher.get('过账人', ''),
                        '违规类型': '调休工作日制单',
                        '日期': date_str,
                        '星期': day_name,
                        '说明': '原本是周末但因调休需要上班'
                    })
                
                # 检查是否为非工作日
                elif not self.is_weekday(date_str_fmt):
                    is_holiday = date_str_fmt in self.all_holidays
                    
                    results.append({
                        '年': voucher.get('年', ''),
                        '月': voucher.get('月', ''),
                        '账簿': voucher.get('账簿', ''),
                        '凭证号': voucher.get('凭证号', ''),
                        '分录号': voucher.get('分录号', ''),
                        '摘要': voucher.get('摘要', ''),
                        '科目': voucher.get('科目', ''),
                        '币种': voucher.get('币种', ''),
                        '借方原币': voucher.get('借方原币', ''),
                        '借方本币': voucher.get('借方本币', ''),
                        '贷方原币': voucher.get('贷方原币', ''),
                        '贷方本币': voucher.get('贷方本币', ''),
                        '来源系统': voucher.get('来源系统', ''),
                        '制单人': voucher.get('制单人', ''),
                        '审核人': voucher.get('审核人', ''),
                        '过账人': voucher.get('过账人', ''),
                        '违规类型': '节假日制单' if is_holiday else '周末制单',
                        '日期': date_str,
                        '星期': day_name
                    })
                    
            except Exception as e:
                continue
        
        return {
            '非工作日制单': results,
            '调休工作日制单': makeup_workday_results
        }
    
    def generate_summary_report(self, year=None):
        """生成汇总报告"""
        print("\n" + "="*80)
        print("📊 会计分录检查汇总报告")
        print("="*80)
        
        # 运行所有检查
        duplicate_check = self.check_duplicate_maker_reviewer()
        unauthorized = self.check_unauthorized_makers()
        empty_desc = self.check_empty_descriptions()
        personnel = self.check_personnel_completeness()
        adjustments = self.check_adjustment_vouchers()
        weekend_vouchers_result = self.check_weekend_vouchers()
        weekend_vouchers = weekend_vouchers_result['非工作日制单']
        makeup_workdays = weekend_vouchers_result['调休工作日制单']
        rpa_compliance = self.check_rpa_reviewer_compliance()
        combination_analysis = self.analyze_maker_reviewer_combinations()
        
        # 创建汇总数据
        summary_data = {
            '检查项目': [
                '制单审核同一人',
                '未授权制单人',
                '无摘要分录',
                '人员信息空值',
                '调整类分录',
                '非工作日制单',
                '调休工作日制单',
                'RPA合规性检查',
                '制单审核组合分析'
            ],
            '违规数量': [
                len(duplicate_check),
                len(unauthorized),
                len(empty_desc),
                len(personnel['空值记录']),
                len(adjustments),
                len(weekend_vouchers),
                len(makeup_workdays),
                len(rpa_compliance['non_compliant_cases']),
                combination_analysis['same_person_count']
            ],
            '状态': [
                '⚠️ 需处理' if duplicate_check else '✅ 正常',
                '⚠️ 需处理' if unauthorized else '✅ 正常',
                '⚠️ 需处理' if empty_desc else '✅ 正常',
                '⚠️ 需处理' if personnel['空值记录'] else '✅ 正常',
                'ℹ️ 需关注' if adjustments else '✅ 无',
                '⚠️ 需处理' if weekend_vouchers else '✅ 正常',
                'ℹ️ 需关注' if makeup_workdays else '✅ 无',
                '⚠️ 需处理' if rpa_compliance['non_compliant_cases'] else '✅ 正常',
                'ℹ️ 统计信息' if combination_analysis['same_person_count'] else '✅ 正常'
            ]
        }
        
        # 创建DataFrame并显示
        summary_df = pd.DataFrame(summary_data)
        print("\n汇总表:")
        print(summary_df.to_string(index=False))
        
        # 确定报告文件名
        if year:
            report_filename = f'会计分录检查报告_{year}年.xlsx'
        else:
            report_filename = '会计分录检查报告.xlsx'
            
        # 保存详细报告到Excel
        try:
            with pd.ExcelWriter(report_filename, engine='openpyxl') as writer:
                # 汇总表
                summary_df.to_excel(writer, sheet_name='汇总报告', index=False)
                
                # 详细检查结果
                if duplicate_check:
                    pd.DataFrame(duplicate_check).to_excel(writer, sheet_name='制单审核同一人', index=False)
                if unauthorized:
                    pd.DataFrame(unauthorized).to_excel(writer, sheet_name='未授权制单人', index=False)
                if empty_desc:
                    pd.DataFrame(empty_desc).to_excel(writer, sheet_name='无摘要分录', index=False)
                if personnel['空值记录']:
                    pd.DataFrame(personnel['空值记录']).to_excel(writer, sheet_name='人员信息空值', index=False)
                if adjustments:
                    pd.DataFrame(adjustments).to_excel(writer, sheet_name='调整类分录', index=False)
                if weekend_vouchers:
                    pd.DataFrame(weekend_vouchers).to_excel(writer, sheet_name='非工作日制单', index=False)
                if makeup_workdays:
                    pd.DataFrame(makeup_workdays).to_excel(writer, sheet_name='调休工作日制单', index=False)
                
                # RPA合规性检查结果
                if rpa_compliance['non_compliant_cases']:
                    pd.DataFrame(rpa_compliance['non_compliant_cases']).to_excel(writer, sheet_name='RPA不合规案例', index=False)
                if rpa_compliance['rpa_analysis_data']:
                    pd.DataFrame(rpa_compliance['rpa_analysis_data']).to_excel(writer, sheet_name='RPA制单人分析', index=False)
                
                # 制单审核组合分析
                if combination_analysis['combination_data']:
                    pd.DataFrame(combination_analysis['combination_data']).to_excel(writer, sheet_name='制单审核组合统计', index=False)
                if combination_analysis['maker_data']:
                    pd.DataFrame(combination_analysis['maker_data']).to_excel(writer, sheet_name='制单人统计', index=False)
                if combination_analysis['reviewer_data']:
                    pd.DataFrame(combination_analysis['reviewer_data']).to_excel(writer, sheet_name='审核人统计', index=False)
                if combination_analysis['same_person_data']:
                    pd.DataFrame(combination_analysis['same_person_data']).to_excel(writer, sheet_name='同一人制单审核统计', index=False)
                
                # 所有人员列表
                personnel_df = pd.DataFrame({
                    '角色': ['制单人', '审核人', '过账人'],
                    '人员名单': [
                        ', '.join(personnel['所有人员']['制单人']),
                        ', '.join(personnel['所有人员']['审核人']),
                        ', '.join(personnel['所有人员']['过账人'])
                    ]
                })
                personnel_df.to_excel(writer, sheet_name='所有人员', index=False)
                
            print(f"\n✅ 详细报告已保存到: {report_filename}")
            
        except Exception as e:
            print(f"❌ 保存报告时出错: {str(e)}")
    
    def run_analysis(self):
        """运行所有检查"""
        print("🔍 开始分析会计分录文件...")
        
        if not self.parse_excel_data():
            return
        
        print(f"📊 共解析到 {len(self.vouchers)} 条分录")
        
        # 1. 检查制单审核同一人
        print("\n" + "="*60)
        print("1️⃣ 检查制单人和审核人为同一人的情况")
        duplicate_check = self.check_duplicate_maker_reviewer()
        if duplicate_check:
            print(f"⚠️  发现 {len(duplicate_check)} 条违规记录:")
            for item in duplicate_check[:10]:  # 只显示前10条
                print(f"   📋 凭证{item.get('凭证号', '')} - {item.get('日期', '')} - 人员:{item.get('制单人', '')}")
            if len(duplicate_check) > 10:
                print(f"   ... 还有 {len(duplicate_check)-10} 条记录")
        else:
            print("✅ 未发现制单审核同一人的情况")
        
        # 2. 检查未授权制单人
        print("\n" + "="*60)
        print("2️⃣ 检查未授权制单人")
        unauthorized = self.check_unauthorized_makers()
        if unauthorized:
            print(f"⚠️  发现 {len(unauthorized)} 条违规记录:")
            for item in unauthorized[:10]:
                violation_type = self.safe_get_field(item, '违规类型')
                voucher_no = self.safe_get_field(item, '凭证号')
                maker = self.safe_get_field(item, '制单人')
                print(f"   📋 {violation_type} - 凭证{voucher_no} - {maker}")
            if len(unauthorized) > 10:
                print(f"   ... 还有 {len(unauthorized)-10} 条记录")
        else:
            print("✅ 未发现未授权制单人")
        
        # 3. 检查无摘要分录
        print("\n" + "="*60)
        print("3️⃣ 检查没有摘要的分录")
        empty_desc = self.check_empty_descriptions()
        if empty_desc:
            print(f"⚠️  发现 {len(empty_desc)} 条无摘要分录")
            for item in empty_desc[:5]:
                print(f"   📋 凭证{item.get('凭证号', '')} - {item.get('日期', '')} - 制单:{item.get('制单人', '')}")
            if len(empty_desc) > 5:
                print(f"   ... 还有 {len(empty_desc)-5} 条记录")
        else:
            print("✅ 所有分录都有摘要")
        
        # 4. 检查人员完整性和空值
        print("\n" + "="*60)
        print("4️⃣ 检查人员完整性和空值")
        personnel = self.check_personnel_completeness()
        
        print("📋 所有制单人员:")
        makers = personnel['所有人员']['制单人']
        print(f"   共{len(makers)}人: {', '.join(makers[:10])}")
        if len(makers) > 10:
            print(f"   ... 还有{len(makers)-10}人")
        
        empty_fields = personnel['空值记录']
        if empty_fields:
            print(f"⚠️  发现 {len(empty_fields)} 个空值字段")
            for item in empty_fields[:5]:
                field = self.safe_get_field(item, '空值字段', '字段')
                voucher_no = self.safe_get_field(item, '凭证号', '凭证')
                print(f"   📋 {field}为空 - 凭证{voucher_no}")
            if len(empty_fields) > 5:
                print(f"   ... 还有 {len(empty_fields)-5} 个空值")
        else:
            print("✅ 所有人员字段均已填写")
        
        # 5. 检查调整分录
        print("\n" + "="*60)
        print("5️⃣ 检查调整类分录")
        adjustments = self.check_adjustment_vouchers()
        if adjustments:
            print(f"📊 发现 {len(adjustments)} 条调整分录，详细信息已记录在Excel报告中")
        else:
            print("✅ 未发现调整类分录")
        
        # 6. 检查非工作日制单
        print("\n" + "="*60)
        print("6️⃣ 检查非工作日制单的凭证")
        weekend_vouchers_result = self.check_weekend_vouchers()
        weekend_vouchers = weekend_vouchers_result['非工作日制单']
        makeup_workdays = weekend_vouchers_result['调休工作日制单']
        
        if weekend_vouchers:
            print(f"⚠️  发现 {len(weekend_vouchers)} 条非工作日制单")
            for item in weekend_vouchers[:10]:
                print(f"   📋 凭证{item.get('凭证号', '')} - {item.get('日期', '')} ({item.get('星期', '')}) - {item.get('类型', '')}")
            if len(weekend_vouchers) > 10:
                print(f"   ... 还有 {len(weekend_vouchers)-10} 条记录")
        else:
            print("✅ 无非工作日制单情况")
            
        # 7. 检查调休工作日制单
        print("\n" + "="*60)
        print("7️⃣ 检查调休工作日制单的凭证")
        if makeup_workdays:
            print(f"ℹ️  发现 {len(makeup_workdays)} 条调休工作日制单")
            for item in makeup_workdays[:10]:
                print(f"   📋 凭证{item.get('凭证号', '')} - {item.get('日期', '')} ({item.get('星期', '')}) - {item.get('说明', '')}")
            if len(makeup_workdays) > 10:
                print(f"   ... 还有 {len(makeup_workdays)-10} 条记录")
            print("   💡 提示: 调休工作日制单属于正常情况，但需要关注是否符合公司政策")
        else:
            print("✅ 无调休工作日制单情况")
        
        # 8. 检查RPA制单人合规性
        print("\n" + "="*60)
        print("8️⃣ 检查RPA制单人合规性")
        rpa_compliance = self.check_rpa_reviewer_compliance()
        if rpa_compliance['rpa_makers']:
            print(f"📊 发现 {len(rpa_compliance['rpa_makers'])} 个RPA制单人")
            print(f"   RPA制单人: {', '.join(rpa_compliance['rpa_makers'])}")
            
            if rpa_compliance['non_compliant_cases']:
                print(f"⚠️  发现 {len(rpa_compliance['non_compliant_cases'])} 条RPA不合规案例")
                for case in rpa_compliance['non_compliant_cases'][:5]:
                    print(f"   📋 {case.get('制单人', '')} → {case.get('审核人', '')} (凭证{case.get('凭证号', '')})")
                if len(rpa_compliance['non_compliant_cases']) > 5:
                    print(f"   ... 还有 {len(rpa_compliance['non_compliant_cases'])-5} 条记录")
            else:
                print("✅ RPA制单人合规性检查通过")
        else:
            print("✅ 未发现RPA制单人")
        
        # 9. 制单审核组合分析
        print("\n" + "="*60)
        print("9️⃣ 制单审核组合分析")
        combination_analysis = self.analyze_maker_reviewer_combinations()
        print(f"📊 组合统计:")
        print(f"   总组合数: {combination_analysis['total_combinations']}")
        print(f"   唯一组合数: {combination_analysis['unique_combinations']}")
        print(f"   制单人数: {combination_analysis['unique_makers']}")
        print(f"   审核人数: {combination_analysis['unique_reviewers']}")
        print(f"   制单审核同一人次数: {combination_analysis['same_person_count']}")
        
        if combination_analysis['top_combinations']:
            print(f"\n📈 前5个最常见组合:")
            for combo, count in combination_analysis['top_combinations'][:5]:
                print(f"   {combo}: {count}次")
        
        # 生成汇总报告
        year = self.get_data_year()
        self.generate_summary_report(year)
        
        print("\n" + "="*60)
        if year:
            print(f"📊 分析完成！报告已保存到 会计分录检查报告_{year}年.xlsx")
        else:
            print("📊 分析完成！报告已保存到 会计分录检查报告.xlsx")
        print("   包含以下检查内容:")
        print("   ✓ 制单审核同一人检查")
        print("   ✓ 未授权制单人检查")
        print("   ✓ 无摘要分录检查")
        print("   ✓ 人员信息完整性检查")
        print("   ✓ 调整类分录检查")
        print("   ✓ 非工作日制单检查")
        print("   ✓ 调休工作日制单检查")
        print("   ✓ RPA制单人合规性检查")
        print("   ✓ 制单审核组合分析")
        print("="*60)
    
    def get_data_year(self):
        """获取数据的年份"""
        if not self.vouchers:
            return None
        
        # 统计各年份的分录数量
        year_counts = {}
        for voucher in self.vouchers:
            date_str = voucher.get('日期', '')
            if date_str and len(date_str) >= 4:
                year = date_str[:4]
                year_counts[year] = year_counts.get(year, 0) + 1
        
        # 返回分录数量最多的年份
        if year_counts:
            return max(year_counts.items(), key=lambda x: x[1])[0]
        return None
    
    def filter_vouchers_by_year(self, year):
        """按年份过滤分录数据"""
        if not year:
            return self.vouchers
        
        filtered_vouchers = []
        for voucher in self.vouchers:
            date_str = voucher.get('日期', '')
            if date_str and len(date_str) >= 4 and date_str[:4] == year:
                filtered_vouchers.append(voucher)
        
        return filtered_vouchers
    
    def run_analysis_by_year(self, year):
        """按年份运行分析"""
        print(f"🔍 开始分析{year}年会计分录文件...")
        
        # 备份原始数据
        original_vouchers = self.vouchers.copy()
        
        # 过滤指定年份的数据
        self.vouchers = self.filter_vouchers_by_year(year)
        
        if not self.vouchers:
            print(f"❌ 没有找到{year}年的分录数据")
            self.vouchers = original_vouchers
            return
        
        print(f"📊 {year}年共有 {len(self.vouchers)} 条分录")
        
        # 运行分析
        self.run_analysis_internal(year)
        
        # 恢复原始数据
        self.vouchers = original_vouchers
    
    def run_analysis_internal(self, year=None):
        """内部分析方法"""
        # 1. 检查制单审核同一人
        print("\n" + "="*60)
        print("1️⃣ 检查制单人和审核人为同一人的情况")
        duplicate_check = self.check_duplicate_maker_reviewer()
        if duplicate_check:
            print(f"⚠️  发现 {len(duplicate_check)} 条违规记录:")
            for item in duplicate_check[:10]:  # 只显示前10条
                voucher_no = self.safe_get_field(item, '凭证号')
                date = self.safe_get_field(item, '日期')
                maker = self.safe_get_field(item, '制单人')
                print(f"   📋 凭证{voucher_no} - {date} - 人员:{maker}")
            if len(duplicate_check) > 10:
                print(f"   ... 还有 {len(duplicate_check)-10} 条记录")
        else:
            print("✅ 未发现制单审核同一人的情况")
        
        # 2. 检查未授权制单人
        print("\n" + "="*60)
        print("2️⃣ 检查未授权制单人")
        unauthorized = self.check_unauthorized_makers()
        if unauthorized:
            print(f"⚠️  发现 {len(unauthorized)} 条违规记录:")
            for item in unauthorized[:10]:
                violation_type = self.safe_get_field(item, '违规类型')
                voucher_no = self.safe_get_field(item, '凭证号')
                maker = self.safe_get_field(item, '制单人')
                print(f"   📋 {violation_type} - 凭证{voucher_no} - {maker}")
            if len(unauthorized) > 10:
                print(f"   ... 还有 {len(unauthorized)-10} 条记录")
        else:
            print("✅ 未发现未授权制单人")
        
        # 3. 检查无摘要分录
        print("\n" + "="*60)
        print("3️⃣ 检查没有摘要的分录")
        empty_desc = self.check_empty_descriptions()
        if empty_desc:
            print(f"⚠️  发现 {len(empty_desc)} 条无摘要分录")
            for item in empty_desc[:5]:
                voucher_no = self.safe_get_field(item, '凭证号')
                date = self.safe_get_field(item, '日期')
                maker = self.safe_get_field(item, '制单人')
                print(f"   📋 凭证{voucher_no} - {date} - 制单:{maker}")
            if len(empty_desc) > 5:
                print(f"   ... 还有 {len(empty_desc)-5} 条记录")
        else:
            print("✅ 所有分录都有摘要")
        
        # 4. 检查人员完整性和空值
        print("\n" + "="*60)
        print("4️⃣ 检查人员完整性和空值")
        personnel = self.check_personnel_completeness()
        
        print("📋 所有制单人员:")
        makers = personnel['所有人员']['制单人']
        print(f"   共{len(makers)}人: {', '.join(makers[:10])}")
        if len(makers) > 10:
            print(f"   ... 还有{len(makers)-10}人")
        
        empty_fields = personnel['空值记录']
        if empty_fields:
            print(f"⚠️  发现 {len(empty_fields)} 个空值字段")
            for item in empty_fields[:5]:
                field = self.safe_get_field(item, '空值字段', '字段')
                voucher_no = self.safe_get_field(item, '凭证号', '凭证')
                print(f"   📋 {field}为空 - 凭证{voucher_no}")
            if len(empty_fields) > 5:
                print(f"   ... 还有 {len(empty_fields)-5} 个空值")
        else:
            print("✅ 所有人员字段均已填写")
        
        # 5. 检查调整分录
        print("\n" + "="*60)
        print("5️⃣ 检查调整类分录")
        adjustments = self.check_adjustment_vouchers()
        if adjustments:
            print(f"📊 发现 {len(adjustments)} 条调整分录，详细信息已记录在Excel报告中")
        else:
            print("✅ 未发现调整类分录")
        
        # 6. 检查非工作日制单
        print("\n" + "="*60)
        print("6️⃣ 检查非工作日制单的凭证")
        weekend_vouchers_result = self.check_weekend_vouchers()
        weekend_vouchers = weekend_vouchers_result['非工作日制单']
        makeup_workdays = weekend_vouchers_result['调休工作日制单']
        
        if weekend_vouchers:
            print(f"⚠️  发现 {len(weekend_vouchers)} 条非工作日制单")
            for item in weekend_vouchers[:10]:
                voucher_no = self.safe_get_field(item, '凭证号')
                date = self.safe_get_field(item, '日期')
                weekday = self.safe_get_field(item, '星期')
                violation_type = self.safe_get_field(item, '违规类型', '类型')
                print(f"   📋 凭证{voucher_no} - {date} ({weekday}) - {violation_type}")
            if len(weekend_vouchers) > 10:
                print(f"   ... 还有 {len(weekend_vouchers)-10} 条记录")
        else:
            print("✅ 无非工作日制单情况")
            
        # 7. 检查调休工作日制单
        print("\n" + "="*60)
        print("7️⃣ 检查调休工作日制单的凭证")
        if makeup_workdays:
            print(f"ℹ️  发现 {len(makeup_workdays)} 条调休工作日制单")
            for item in makeup_workdays[:10]:
                voucher_no = self.safe_get_field(item, '凭证号')
                date = self.safe_get_field(item, '日期')
                weekday = self.safe_get_field(item, '星期')
                description = self.safe_get_field(item, '说明')
                print(f"   📋 凭证{voucher_no} - {date} ({weekday}) - {description}")
            if len(makeup_workdays) > 10:
                print(f"   ... 还有 {len(makeup_workdays)-10} 条记录")
            print("   💡 提示: 调休工作日制单属于正常情况，但需要关注是否符合公司政策")
        else:
            print("✅ 无调休工作日制单情况")
        
        # 8. 检查RPA制单人合规性
        print("\n" + "="*60)
        print("8️⃣ 检查RPA制单人合规性")
        rpa_compliance = self.check_rpa_reviewer_compliance()
        if rpa_compliance['rpa_makers']:
            print(f"📊 发现 {len(rpa_compliance['rpa_makers'])} 个RPA制单人")
            print(f"   RPA制单人: {', '.join(rpa_compliance['rpa_makers'])}")
            
            if rpa_compliance['non_compliant_cases']:
                print(f"⚠️  发现 {len(rpa_compliance['non_compliant_cases'])} 条RPA不合规案例")
                for case in rpa_compliance['non_compliant_cases'][:5]:
                    maker = self.safe_get_field(case, '制单人')
                    reviewer = self.safe_get_field(case, '审核人')
                    voucher_no = self.safe_get_field(case, '凭证号')
                    print(f"   📋 {maker} → {reviewer} (凭证{voucher_no})")
                if len(rpa_compliance['non_compliant_cases']) > 5:
                    print(f"   ... 还有 {len(rpa_compliance['non_compliant_cases'])-5} 条记录")
            else:
                print("✅ RPA制单人合规性检查通过")
        else:
            print("✅ 未发现RPA制单人")
        
        # 9. 制单审核组合分析
        print("\n" + "="*60)
        print("9️⃣ 制单审核组合分析")
        combination_analysis = self.analyze_maker_reviewer_combinations()
        print(f"📊 组合统计:")
        print(f"   总组合数: {combination_analysis['total_combinations']}")
        print(f"   唯一组合数: {combination_analysis['unique_combinations']}")
        print(f"   制单人数: {combination_analysis['unique_makers']}")
        print(f"   审核人数: {combination_analysis['unique_reviewers']}")
        print(f"   制单审核同一人次数: {combination_analysis['same_person_count']}")
        
        if combination_analysis['top_combinations']:
            print(f"\n📈 前5个最常见组合:")
            for combo, count in combination_analysis['top_combinations'][:5]:
                print(f"   {combo}: {count}次")
        
        # 生成汇总报告
        self.generate_summary_report(year)
        
        print("\n" + "="*60)
        if year:
            print(f"📊 {year}年分析完成！报告已保存到 会计分录检查报告_{year}年.xlsx")
        else:
            print("📊 分析完成！报告已保存到 会计分录检查报告.xlsx")
        print("   包含以下检查内容:")
        print("   ✓ 制单审核同一人检查")
        print("   ✓ 未授权制单人检查")
        print("   ✓ 无摘要分录检查")
        print("   ✓ 人员信息完整性检查")
        print("   ✓ 调整类分录检查")
        print("   ✓ 非工作日制单检查")
        print("   ✓ 调休工作日制单检查")
        print("   ✓ RPA制单人合规性检查")
        print("   ✓ 制单审核组合分析")
        print("="*60)

def main():
    """主函数"""
    import glob
    import sys
    
    # 检查命令行参数
    target_year = None
    if len(sys.argv) > 1:
        arg = sys.argv[1].strip()
        if arg in ['2022', '2023', '2024', '2025']:
            target_year = arg
            print(f"🎯 通过命令行参数指定分析{target_year}年数据")
        elif arg == 'all':
            target_year = 'all'
            print("🎯 通过命令行参数指定分析所有年份数据")
        else:
            print(f"❌ 无效的命令行参数: {arg}")
            print("💡 有效参数: 2022, 2023, 2024, 2025, all")
            print("💡 示例: python accounting_voucher_analyzer_2025.py 2025")
            sys.exit(1)
    
    # 自动查找当前目录下所有2022-2025年的JE文件
    current_dir = "d:\\User Data\\yangfan15\\Desktop\\testing"
    
    # 查找所有可能的2022-2025年JE文件模式
    years = ['2022', '2023', '2024', '2025']
    all_patterns = []
    
    for year in years:
        patterns = [
            os.path.join(current_dir, f"{year}je*.xlsx"),
            os.path.join(current_dir, f"{year}JE*.xlsx"),
            os.path.join(current_dir, f"*{year}*je*.xlsx"),
            os.path.join(current_dir, f"*{year}*JE*.xlsx")
        ]
        all_patterns.extend(patterns)
    
    print("🚀 启动2022-2025年会计分录检查分析...")
    print("🔍 正在搜索2022-2025年JE文件...")
    
    # 收集所有匹配的文件
    all_je_files = set()
    for pattern in all_patterns:
        files = glob.glob(pattern)
        all_je_files.update(files)
    
    # 转换为列表并按年份和文件名排序
    all_je_files = sorted(list(all_je_files))
    
    # 按年份分组显示
    files_by_year = {year: [] for year in years}
    for file_path in all_je_files:
        filename = os.path.basename(file_path)
        for year in years:
            if year in filename:
                files_by_year[year].append(file_path)
                break
    
    total_files = len(all_je_files)
    print(f"📁 找到 {total_files} 个JE文件:")
    
    # 显示按年份分组的文件
    for year in years:
        year_files = files_by_year[year]
        if year_files:
            print(f"\n📅 {year}年 ({len(year_files)}个文件):")
            for file_path in year_files:
                print(f"   ✓ {os.path.basename(file_path)}")
    
    # 检查文件是否存在
    existing_files = []
    for file_path in all_je_files:
        if os.path.exists(file_path):
            existing_files.append(file_path)
        else:
            print(f"   ❌ {os.path.basename(file_path)} - 文件不存在")
    
    if not existing_files:
        print("❌ 没有找到任何JE文件")
        print("💡 请确保文件名包含年份(2022-2025)和'je'或'JE'")
        sys.exit(1)
    
    # 根据命令行参数决定要处理的文件
    files_to_process = existing_files
    if target_year and target_year != 'all':
        # 如果指定了具体年份，只处理该年份的文件
        files_to_process = []
        for file_path in existing_files:
            filename = os.path.basename(file_path)
            if target_year in filename:
                files_to_process.append(file_path)
        
        if not files_to_process:
            print(f"❌ 没有找到{target_year}年的JE文件")
            sys.exit(1)
        
        print(f"🎯 仅处理{target_year}年的 {len(files_to_process)} 个文件")
    
    # 创建分析器并处理文件
    analyzer = AccountingVoucherAnalyzer()
    
    print(f"\n🔄 开始处理 {len(files_to_process)} 个文件...")
    
    # 逐个处理文件
    processed_count = 0
    for i, file_path in enumerate(files_to_process, 1):
        print(f"\n📊 正在处理第 {i}/{len(files_to_process)} 个文件: {os.path.basename(file_path)}")
        analyzer.file_path = file_path
        if analyzer.parse_excel_data():
            processed_count += 1
            print(f"✅ 文件 {os.path.basename(file_path)} 处理完成")
        else:
            print(f"❌ 处理文件 {os.path.basename(file_path)} 失败")
    
    if analyzer.vouchers:
        print(f"\n📊 成功处理 {processed_count}/{len(files_to_process)} 个文件")
        print(f"📊 总计处理了 {len(analyzer.vouchers)} 条分录")
        
        # 按年份统计分录数量
        vouchers_by_year = {}
        for voucher in analyzer.vouchers:
            date_str = voucher.get('日期', '')
            if date_str and len(date_str) >= 4:
                year = date_str[:4]
                vouchers_by_year[year] = vouchers_by_year.get(year, 0) + 1
        
        print("\n📈 各年份分录统计:")
        available_years = []
        for year in sorted(vouchers_by_year.keys()):
            print(f"   {year}年: {vouchers_by_year[year]:,} 条分录")
            available_years.append(year)
        
        # 根据命令行参数或用户选择执行分析
        if target_year:
            # 通过命令行参数指定
            if target_year == 'all':
                print("\n🔄 开始合并分析所有年份数据...")
                analyzer.run_analysis()
            elif target_year in ['2022', '2023', '2024', '2025']:
                # 直接分析指定年份（此时已经只加载了该年份的数据）
                print(f"\n🔄 开始分析{target_year}年数据...")
                analyzer.run_analysis_internal(target_year)
            else:
                print(f"❌ 指定的年份{target_year}在数据中不存在")
                print(f"📊 可用年份: {', '.join(available_years)}")
                print("   默认进行合并分析...")
                analyzer.run_analysis()
        else:
            # 交互式选择
            print("\n" + "="*60)
            print("🎯 请选择分析方式:")
            print("   0. 合并所有年份数据进行分析")
            for i, year in enumerate(available_years, 1):
                print(f"   {i}. 仅分析{year}年数据")
            print("="*60)
            print("\n💡 提示: 请在控制台中输入选择数字")
            print("💡 或者使用命令行参数: python accounting_voucher_analyzer_2025.py [2022|2023|2024|2025|all]")
            
            try:
                # 确保输入提示清晰可见
                choice_input = input(f"\n请输入选择 (0-{len(available_years)}): ")
                print(f"\n📝 您的选择: {choice_input}")
                
                choice = int(choice_input.strip())
                
                if choice == 0:
                    # 合并分析
                    print("\n🔄 开始合并分析所有年份数据...")
                    analyzer.run_analysis()
                elif 1 <= choice <= len(available_years):
                    # 按年分析
                    selected_year = available_years[choice - 1]
                    print(f"\n🔄 开始分析{selected_year}年数据...")
                    analyzer.run_analysis_by_year(selected_year)
                else:
                    print(f"❌ 无效选择 '{choice}'，默认进行合并分析")
                    analyzer.run_analysis()
            except (ValueError, KeyboardInterrupt, EOFError) as e:
                print(f"\n❌ 输入处理异常: {type(e).__name__}")
                print("   可能原因: 在非交互式环境中运行或输入被中断")
                print("   💡 建议使用命令行参数: python accounting_voucher_analyzer_2025.py all")
                print("   默认进行合并分析...")
                analyzer.run_analysis()
            except Exception as e:
                print(f"\n❌ 未知异常: {e}")
                print("   默认进行合并分析...")
                analyzer.run_analysis()
    else:
        print("❌ 没有成功处理任何数据")

if __name__ == "__main__":
    main()