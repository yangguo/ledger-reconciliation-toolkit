import pandas as pd
import numpy as np
from datetime import datetime
import json
import os
import argparse
from typing import List, Dict, Optional, Tuple
from utils import parse_currency_value, extract_account_code
class ReconciliationConfig:
    """
    对账配置类，用于管理对账参数
    """
    def __init__(self, config_file: Optional[str] = None):
        self.config = self._load_default_config()
        if config_file and os.path.exists(config_file):
            self._load_config_file(config_file)
    
    def _load_default_config(self) -> Dict:
        """加载默认配置"""
        return {
            "target_patterns": [],  # 账套筛选模式列表
            "je_files": [],  # JE文件列表
            "tb_file": "",  # TB文件
            "output_prefix": "对账报告",  # 输出文件前缀
            "threshold": 0.01,  # 差异阈值
            "je_columns": {  # JE文件列名映射
                "book": "账簿",
                "subject": "科目",
                "debit": "借方本币",
                "credit": "贷方本币"
            },
            "tb_columns": {  # TB文件列名映射
                "book": ["核算账簿名称", "主体账簿", "账簿"],
                "account_code": "科目编码",
                "debit": ["本期借方.1", "本期借方发生.1", "本期借方", "借方累计.1", "借方累计"],
                "credit": ["本期贷方.1", "本期贷方发生.1", "本期贷方", "贷方累计.1", "贷方累计"],
                "debit_col_index": None,  # 借方列索引（用于处理重复列名）
                "credit_col_index": None  # 贷方列索引（用于处理重复列名）
            },
            "header_row_index": None,  # TB文件表头行索引
            "default_book": "默认账簿",  # 当未找到账簿列时的默认账簿名
            "summary_patterns": [],  # 额外的汇总记录过滤模式
            "filter_invalid_codes": ["总计", "核算账簿累计", "合计", "nan", "币种累计", "科目编码"],
            "filter_patterns": ["币种累计", "核算单位", "制单人", "打印时间"]
        }
    
    def _load_config_file(self, config_file: str):
        """从文件加载配置"""
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                file_config = json.load(f)
                self.config.update(file_config)
        except Exception as e:
            print(f"警告: 无法加载配置文件 {config_file}: {e}")
    
    def set_target_patterns(self, patterns: List[str]):
        """设置账套筛选模式"""
        self.config["target_patterns"] = patterns
    
    def set_files(self, je_files: List[str], tb_file: str):
        """设置文件路径"""
        self.config["je_files"] = je_files
        self.config["tb_file"] = tb_file
    
    def get(self, key: str, default=None):
        """获取配置值"""
        return self.config.get(key, default)

class GeneralReconciliationAnalyzer:
    """
    通用对账分析器
    """
    def __init__(self, config: ReconciliationConfig):
        self.config = config
        self.je_data = None
        self.tb_data = None
    
    def load_je_files(self) -> pd.DataFrame:
        """
        加载并合并JE文件
        """
        je_files = self.config.get("je_files", [])
        if not je_files:
            raise ValueError("未指定JE文件")
        
        print(f"正在加载 {len(je_files)} 个JE文件...")
        
        all_dfs = []
        for file_path in je_files:
            if not os.path.exists(file_path):
                print(f"警告: JE文件不存在: {file_path}")
                continue
            
            try:
                df = pd.read_excel(file_path)
                print(f"加载JE文件 {file_path}: {len(df)} 行")
                all_dfs.append(df)
            except Exception as e:
                print(f"错误: 无法加载JE文件 {file_path}: {e}")
        
        if not all_dfs:
            raise ValueError("没有成功加载任何JE文件")
        
        # 检查列名一致性
        if len(all_dfs) > 1:
            all_columns = [list(df.columns) for df in all_dfs]
            if not all(cols == all_columns[0] for cols in all_columns):
                print("警告: JE文件的列名不一致，将使用共同列")
                common_cols = set(all_dfs[0].columns)
                for df in all_dfs[1:]:
                    common_cols &= set(df.columns)
                common_cols = list(common_cols)
                print(f"共同列: {common_cols}")
                all_dfs = [df[common_cols] for df in all_dfs]
        
        # 合并数据
        self.je_data = pd.concat(all_dfs, ignore_index=True)
        print(f"合并后JE数据: {len(self.je_data)} 行")
        
        return self.je_data
    
    def load_tb_file(self) -> pd.DataFrame:
        """
        加载TB文件
        """
        tb_file = self.config.get("tb_file")
        if not tb_file:
            raise ValueError("未指定TB文件")
        
        if not os.path.exists(tb_file):
            raise FileNotFoundError(f"TB文件不存在: {tb_file}")
        
        print(f"正在加载TB文件: {tb_file}")
        self.tb_data = pd.read_excel(tb_file)
        print(f"TB数据: {len(self.tb_data)} 行")
        
        return self.tb_data
    
    def _detect_header_row(self, tb_df: pd.DataFrame) -> int:
        """
        自动检测TB文件的表头行位置
        """
        header_row_config = self.config.get("header_row_index", None)
        if header_row_config is not None:
            print(f"使用配置指定的表头行: {header_row_config}")
            return header_row_config
        
        # 自动检测表头行
        for i in range(min(10, len(tb_df))):
            row_values = tb_df.iloc[i].astype(str).str.strip()
            # 检查是否包含常见的表头关键词
            header_keywords = ['科目编码', '科目名称', '借方', '贷方', '账簿', '核算账簿']
            if any(keyword in ' '.join(row_values.values) for keyword in header_keywords):
                print(f"自动检测到表头行: 第{i}行")
                return i
        
        print("未检测到表头行，使用第0行")
        return 0
    
    def _detect_tb_columns(self, tb_df: pd.DataFrame) -> Dict[str, str]:
        """
        自动检测TB文件的列名，支持列索引访问
        """
        tb_columns = self.config.get("tb_columns", {})
        detected_columns = {}
        
        # 检测账簿列
        book_candidates = tb_columns.get("book", [])
        for col in tb_df.columns:
            col_str = str(col).strip()
            if any(candidate in col_str for candidate in book_candidates):
                detected_columns["book"] = col
                break
        
        # 检测借方列 - 支持列索引配置
        debit_candidates = tb_columns.get("debit", [])
        debit_col_index = tb_columns.get("debit_col_index", None)
        
        if debit_col_index is not None and debit_col_index < len(tb_df.columns):
            detected_columns["debit"] = tb_df.columns[debit_col_index]
            detected_columns["debit_index"] = debit_col_index
        else:
            for candidate in debit_candidates:
                if candidate in tb_df.columns:
                    detected_columns["debit"] = candidate
                    break
        
        # 检测贷方列 - 支持列索引配置
        credit_candidates = tb_columns.get("credit", [])
        credit_col_index = tb_columns.get("credit_col_index", None)
        
        if credit_col_index is not None and credit_col_index < len(tb_df.columns):
            detected_columns["credit"] = tb_df.columns[credit_col_index]
            detected_columns["credit_index"] = credit_col_index
        else:
            for candidate in credit_candidates:
                if candidate in tb_df.columns:
                    detected_columns["credit"] = candidate
                    break
        
        # 科目编码列
        account_code_col = tb_columns.get("account_code", "科目编码")
        if account_code_col in tb_df.columns:
            detected_columns["account_code"] = account_code_col
        
        print(f"检测到的TB列名: {detected_columns}")
        return detected_columns
    
    def prepare_je_data(self, target_patterns: List[str]) -> pd.DataFrame:
        """
        准备JE数据
        """
        if self.je_data is None:
            raise ValueError("JE数据未加载")
        
        print(f"正在准备JE数据，筛选账套模式: {target_patterns}")
        
        je_clean = self.je_data.copy()
        je_columns = self.config.get("je_columns", {})
        
        # 数据清理
        debit_col = je_columns.get("debit", "借方本币")
        credit_col = je_columns.get("credit", "贷方本币")
        book_col = je_columns.get("book", "账簿")
        subject_col = je_columns.get("subject", "科目")
        
        je_clean[debit_col] = pd.to_numeric(je_clean[debit_col], errors='coerce').fillna(0)
        je_clean[credit_col] = pd.to_numeric(je_clean[credit_col], errors='coerce').fillna(0)
        je_clean[book_col] = je_clean[book_col].astype(str).str.strip()
        
        # 筛选账套
        print(f"JE原始记录数: {len(je_clean)}")
        if target_patterns:
            pattern_filter = pd.Series([False] * len(je_clean))
            for pattern in target_patterns:
                pattern_filter |= je_clean[book_col].str.contains(pattern, na=False)
            je_clean = je_clean[pattern_filter]
            print(f"筛选账套后JE记录数: {len(je_clean)}")
            
            # 显示找到的账套
            found_books = je_clean[book_col].unique()
            print(f"找到的账套数量: {len(found_books)}")
            for i, book in enumerate(found_books):
                print(f"  {i+1}. {book}")
        
        if len(je_clean) == 0:
            print(f"警告: 在JE数据中未找到匹配的账套")
            return pd.DataFrame()
        
        # 提取科目编码
        print("正在提取JE科目编码...")
        je_clean['科目编码'] = je_clean[subject_col].apply(extract_account_code)
        
        # 过滤掉无效的科目编码
        je_clean = je_clean[je_clean['科目编码'] != '']
        
        # 过滤掉借贷均为0的原始记录
        print("正在过滤借贷均为0的JE原始记录...")
        before_zero_filter = len(je_clean)
        je_clean = je_clean[~((je_clean[debit_col] == 0) & (je_clean[credit_col] == 0))]
        after_zero_filter = len(je_clean)
        print(f"过滤借贷均为0原始记录: 过滤前 {before_zero_filter} 条，过滤后 {after_zero_filter} 条，已过滤 {before_zero_filter - after_zero_filter} 条")
        
        # 按账簿和科目编码汇总
        je_summary = je_clean.groupby([book_col, '科目编码']).agg({
            debit_col: 'sum',
            credit_col: 'sum'
        }).reset_index()
        
        # 重命名列
        je_summary = je_summary.rename(columns={
            book_col: '账簿',
            debit_col: '借方本币',
            credit_col: '贷方本币'
        })
        
        # 创建原始记录存在性标记
        je_exists = je_clean.groupby([book_col, '科目编码']).size().reset_index(name='je_record_count')
        je_exists['je_exists'] = True
        je_exists = je_exists.rename(columns={book_col: '账簿'})
        
        # 合并汇总数据和存在性标记
        je_final = je_summary.merge(je_exists, on=['账簿', '科目编码'], how='left')
        
        # 过滤掉汇总后借贷均为0的记录
        print("正在过滤汇总后借贷均为0的JE记录...")
        before_summary_zero_filter = len(je_final)
        threshold = 1e-6
        je_final = je_final[~((abs(je_final['借方本币']) <= threshold) & (abs(je_final['贷方本币']) <= threshold))]
        after_summary_zero_filter = len(je_final)
        print(f"过滤汇总后借贷均为0记录（阈值{threshold}）: 过滤前 {before_summary_zero_filter} 条，过滤后 {after_summary_zero_filter} 条，已过滤 {before_summary_zero_filter - after_summary_zero_filter} 条")
        
        print(f"JE最终记录数: {len(je_final)}")
        return je_final
    
    def prepare_tb_data(self, target_patterns: List[str]) -> pd.DataFrame:
        """
        准备TB数据，支持不同的TB格式
        """
        if self.tb_data is None:
            raise ValueError("TB数据未加载")
        
        print(f"正在准备TB数据，筛选账套模式: {target_patterns}")
        
        tb_raw = self.tb_data.copy()
        print(f"TB原始数据形状: {tb_raw.shape}")
        
        # 检测表头行位置
        header_row_index = self._detect_header_row(tb_raw)
        
        # 如果表头不在第0行，重新构建DataFrame
        if header_row_index > 0:
            new_columns = tb_raw.iloc[header_row_index].tolist()
            tb_clean = tb_raw.iloc[header_row_index + 1:].copy()
            tb_clean.columns = new_columns
            tb_clean = tb_clean.reset_index(drop=True)
            print(f"使用第{header_row_index}行作为表头")
        else:
            tb_clean = tb_raw.copy()
        
        # 处理重复列名问题
        columns = tb_clean.columns.tolist()
        new_columns = []
        column_counts = {}
        
        for col in columns:
            if col in column_counts:
                column_counts[col] += 1
                new_columns.append(f"{col}_{column_counts[col]}")
            else:
                column_counts[col] = 0
                new_columns.append(col)
        
        tb_clean.columns = new_columns
        
        # 清理TB数据
        tb_clean = self._clean_tb_data(tb_clean)
        
        print(f"TB文件列名: {tb_clean.columns.tolist()}")
        print(f"TB处理后记录数: {len(tb_clean)}")
        
        # 自动检测列名
        detected_columns = self._detect_tb_columns(tb_clean)
        
        # 处理货币字段
        print("正在解析TB货币字段...")
        debit_col = detected_columns.get("debit")
        credit_col = detected_columns.get("credit")
        book_col = detected_columns.get("book")
        account_code_col = detected_columns.get("account_code", "科目编码")
        
        if debit_col and debit_col in tb_clean.columns:
            tb_clean[debit_col] = tb_clean[debit_col].apply(parse_currency_value)
        else:
            print(f"警告: 未找到借方列")
            tb_clean['借方本币'] = 0
            debit_col = '借方本币'
            
        if credit_col and credit_col in tb_clean.columns:
            tb_clean[credit_col] = tb_clean[credit_col].apply(parse_currency_value)
        else:
            print(f"警告: 未找到贷方列")
            tb_clean['贷方本币'] = 0
            credit_col = '贷方本币'
        
        # 处理账簿筛选
        if book_col and book_col in tb_clean.columns and target_patterns:
            tb_clean[book_col] = tb_clean[book_col].astype(str).str.strip()
            print(f"筛选前TB记录数: {len(tb_clean)}")
            
            pattern_filter = pd.Series([False] * len(tb_clean))
            for pattern in target_patterns:
                pattern_filter |= tb_clean[book_col].str.contains(pattern, na=False)
            tb_clean = tb_clean[pattern_filter]
            print(f"筛选账套后TB记录数: {len(tb_clean)}")
            
            # 显示找到的账套
            if len(tb_clean) > 0:
                found_books = tb_clean[book_col].unique()
                print(f"TB中找到的账套数量: {len(found_books)}")
                for i, book in enumerate(found_books):
                    print(f"  {i+1}. {book}")
            # 标记TB数据有真实的账簿列
            self._tb_has_real_book_column = True
        elif not book_col:
            print("未找到账簿列，将添加目标账簿")
            # 使用target_patterns中的第一个模式作为账簿名
            target_book = target_patterns[0] if target_patterns else self.config.get("default_book", "默认账簿")
            tb_clean['账簿'] = target_book
            book_col = '账簿'
            # 标记TB数据没有真实的账簿列
            self._tb_has_real_book_column = False
            print(f"已添加目标账簿: {target_book}，将仅按科目编码进行对账")
        
        # 处理科目编码
        if account_code_col in tb_clean.columns:
            tb_clean[account_code_col] = tb_clean[account_code_col].astype(str).str.strip()
        else:
            raise ValueError(f"未找到科目编码列: {account_code_col}")
        
        # 过滤掉汇总记录
        print("正在过滤TB汇总记录...")
        original_count = len(tb_clean)
        
        # 使用新的过滤函数
        tb_clean = self._filter_summary_records(tb_clean, account_code_col)
        
        # 过滤配置中指定的无效代码
        filter_invalid_codes = self.config.get("filter_invalid_codes", [])
        if filter_invalid_codes:
            tb_clean = tb_clean[~tb_clean[account_code_col].isin(filter_invalid_codes)]
        
        # 过滤包含特定模式的记录
        filter_patterns = self.config.get("filter_patterns", [])
        for pattern in filter_patterns:
            tb_clean = tb_clean[~tb_clean[account_code_col].str.contains(pattern, na=False)]
        
        filtered_count = len(tb_clean)
        print(f"过滤前记录数: {original_count}, 过滤后记录数: {filtered_count}, 已过滤汇总记录: {original_count - filtered_count}")
        
        # 重命名列以便匹配
        rename_dict = {
            debit_col: '借方本币',
            credit_col: '贷方本币',
            account_code_col: '科目编码'
        }
        if book_col:
            rename_dict[book_col] = '账簿'
        
        tb_clean = tb_clean.rename(columns=rename_dict)
        
        # 提取科目编码
        print("正在标准化TB科目编码...")
        tb_clean['科目编码'] = tb_clean['科目编码'].apply(extract_account_code)
        
        # 过滤有效记录
        tb_clean = tb_clean.dropna(subset=['账簿', '科目编码'])
        tb_clean = tb_clean[tb_clean['科目编码'] != '']
        
        print(f"TB过滤后记录数: {len(tb_clean)}")
        
        # 显示前几条记录用于验证
        if len(tb_clean) > 0:
            print("\n前5条TB记录:")
            for idx, row in tb_clean.head().iterrows():
                print(f"  科目编码: {row['科目编码']}, 借方: {row['借方本币']}, 贷方: {row['贷方本币']}")
        
        # 按账簿和科目编码汇总
        tb_summary = tb_clean.groupby(['账簿', '科目编码']).agg({
            '借方本币': 'sum',
            '贷方本币': 'sum'
        }).reset_index()
        
        # 创建原始记录存在性标记
        tb_exists = tb_clean.groupby(['账簿', '科目编码']).size().reset_index(name='tb_record_count')
        tb_exists['tb_exists'] = True
        
        # 合并汇总数据和存在性标记
        tb_final = tb_summary.merge(tb_exists, on=['账簿', '科目编码'], how='left')
        
        # 确保数值列是数值类型
        if len(tb_final) > 0:
            tb_final['借方本币'] = pd.to_numeric(tb_final['借方本币'], errors='coerce').fillna(0)
            tb_final['贷方本币'] = pd.to_numeric(tb_final['贷方本币'], errors='coerce').fillna(0)
        
        # 过滤借贷均为0的记录
        print("正在过滤借贷均为0的TB记录...")
        before_zero_filter = len(tb_final)
        threshold = 1e-6
        
        if len(tb_final) > 0:
            zero_amount_mask = (abs(tb_final['借方本币']) <= threshold) & (abs(tb_final['贷方本币']) <= threshold)
            tb_final = tb_final[~zero_amount_mask]
        else:
            zero_amount_mask = pd.Series([], dtype=bool)
        
        after_zero_filter = len(tb_final)
        print(f"过滤借贷均为0记录（阈值{threshold}）: 过滤前 {before_zero_filter} 条，过滤后 {after_zero_filter} 条，已过滤 {zero_amount_mask.sum()} 条")
        
        print(f"TB汇总后记录数: {len(tb_final)}")
        
        return tb_final
    
    def perform_reconciliation(self, je_summary: pd.DataFrame, tb_summary: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
        """
        执行对账分析
        """
        print("\n正在执行对账分析...")
        
        if len(je_summary) == 0 or len(tb_summary) == 0:
            print("警告: JE或TB数据为空，无法执行对账")
            return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()
        
        # 检查TB数据是否有真实的账簿列（非默认添加的）
        tb_has_real_book_column = hasattr(self, '_tb_has_real_book_column') and self._tb_has_real_book_column
        
        if tb_has_real_book_column:
            # TB有真实账簿列，按账簿和科目编码合并
            print("TB数据包含真实账簿列，按账簿和科目编码进行对账")
            merge_on = ['账簿', '科目编码']
        else:
            # TB没有真实账簿列，只按科目编码合并
            print("TB数据未包含真实账簿列，仅按科目编码进行对账")
            merge_on = ['科目编码']
            # 为了保持数据结构一致，给JE和TB数据添加相同的账簿标识
            if '账簿' in je_summary.columns and '账簿' in tb_summary.columns:
                # 统一账簿名称以便合并
                unified_book_name = "统一账簿"
                je_summary = je_summary.copy()
                tb_summary = tb_summary.copy()
                je_summary['账簿'] = unified_book_name
                tb_summary['账簿'] = unified_book_name
                merge_on = ['账簿', '科目编码']
        
        # 合并数据
        merged = pd.merge(
            je_summary, tb_summary, 
            on=merge_on, 
            how='outer', 
            suffixes=('_JE', '_TB')
        )
        
        # 填充缺失值
        merged['借方本币_JE'] = merged['借方本币_JE'].fillna(0)
        merged['贷方本币_JE'] = merged['贷方本币_JE'].fillna(0)
        merged['借方本币_TB'] = merged['借方本币_TB'].fillna(0)
        merged['贷方本币_TB'] = merged['贷方本币_TB'].fillna(0)
        merged['je_exists'] = merged['je_exists'].fillna(False)
        merged['tb_exists'] = merged['tb_exists'].fillna(False)
        
        # 计算差异
        merged['借方差异'] = merged['借方本币_JE'] - merged['借方本币_TB']
        merged['贷方差异'] = merged['贷方本币_JE'] - merged['贷方本币_TB']
        
        # 设置差异阈值
        threshold = self.config.get("threshold", 0.01)
        
        # 分类逻辑：基于原始记录的存在性
        both_exist = merged['je_exists'] & merged['tb_exists']
        je_only = merged['je_exists'] & ~merged['tb_exists']
        tb_only = ~merged['je_exists'] & merged['tb_exists']
        
        # 两者都有但存在差异
        both_diff = merged[both_exist & 
                          ((abs(merged['借方差异']) > threshold) | 
                           (abs(merged['贷方差异']) > threshold))].copy()
        
        # 两者都有且无差异（在阈值范围内）
        both_no_diff = merged[both_exist & 
                             (abs(merged['借方差异']) <= threshold) & 
                             (abs(merged['贷方差异']) <= threshold)].copy()
        
        # 仅在JE中
        je_only_records = merged[je_only].copy()
        
        # 仅在TB中
        tb_only_records = merged[tb_only].copy()
        
        print(f"\n=== 对账结果 ===")
        print(f"两者都有但存在差异: {len(both_diff)} 条")
        print(f"两者都有且无差异: {len(both_no_diff)} 条")
        print(f"仅在JE中: {len(je_only_records)} 条")
        print(f"仅在TB中: {len(tb_only_records)} 条")
        print(f"总差异记录数: {len(both_diff) + len(je_only_records) + len(tb_only_records)}")
        print(f"总记录数（含无差异）: {len(both_diff) + len(both_no_diff) + len(je_only_records) + len(tb_only_records)}")
        
        return both_diff, both_no_diff, je_only_records, tb_only_records, merged
    
    def _clean_tb_data(self, tb_df: pd.DataFrame) -> pd.DataFrame:
        """
        清理TB数据，移除无效行和格式化数据
        """
        # 移除完全空白的行
        tb_df = tb_df.dropna(how='all')
        
        # 移除所有列都为空字符串的行
        string_cols = tb_df.select_dtypes(include=['object']).columns
        if len(string_cols) > 0:
            mask = ~(tb_df[string_cols].astype(str).eq('').all(axis=1))
            tb_df = tb_df[mask]
        
        return tb_df.reset_index(drop=True)
    
    def _filter_summary_records(self, tb_df: pd.DataFrame, account_code_col: str) -> pd.DataFrame:
        """
        过滤TB中的汇总记录
        """
        # 基本的汇总记录过滤
        summary_keywords = ['合计', '小计', '总计', '汇总']
        mask = ~tb_df[account_code_col].astype(str).str.contains('|'.join(summary_keywords), na=False)
        
        # 配置中的额外过滤模式
        summary_patterns = self.config.get("summary_patterns", [])
        for pattern in summary_patterns:
            mask &= ~tb_df[account_code_col].astype(str).str.contains(pattern, na=False)
        
        return tb_df[mask]
    
    def check_je_voucher_gaps(self, target_patterns: List[str]) -> Tuple[pd.DataFrame, Dict]:
        """
        检测JE凭证号跳号情况
        检查相同年、月、账簿、凭证类型下凭证号是否连续编号
        支持带前缀的凭证号格式（如：财字凭证-1, 报销字凭证-1）
        """
        print(f"\n正在检测JE凭证号跳号情况（筛选模式: {target_patterns}）...")
        
        if self.je_data is None:
            print("警告: JE数据未加载")
            return pd.DataFrame(), {}
        
        # 筛选目标账套数据
        if target_patterns:
            mask = pd.Series([False] * len(self.je_data))
            for pattern in target_patterns:
                mask |= self.je_data['账簿'].astype(str).str.contains(pattern, na=False)
            je_clean = self.je_data[mask].copy()
        else:
            je_clean = self.je_data.copy()
        
        if len(je_clean) == 0:
            print(f"警告: 在JE数据中未找到包含'{target_patterns}'的记录")
            return pd.DataFrame(), {}
        
        # 只保留必要的列以减少内存使用
        required_cols = ['年', '月', '账簿', '凭证号']
        available_cols = [col for col in required_cols if col in je_clean.columns]
        
        if '凭证号' not in available_cols:
            print("警告: JE数据中缺少'凭证号'列")
            return pd.DataFrame(), {}
        
        je_clean = je_clean[available_cols].copy()
        
        # 清理年份数据
        if '年' in je_clean.columns:
            try:
                je_clean['年'] = pd.to_numeric(je_clean['年'], errors='coerce')
                je_clean = je_clean.dropna(subset=['年'])
                je_clean['年'] = je_clean['年'].astype(int)
            except Exception as e:
                print(f"警告: 年份数据清理失败: {e}")
                return pd.DataFrame(), {}
        
        # 清理月份数据，保持原始月份值（包括'12A'等）
        if '月' in je_clean.columns:
            try:
                # 将月份转换为字符串并清理空白字符
                je_clean['月'] = je_clean['月'].astype(str).str.strip()
                
                # 过滤掉空值或无效的月份
                valid_month_mask = (je_clean['月'].notna()) & (je_clean['月'] != '') & (je_clean['月'] != 'nan')
                before_month_filter = len(je_clean)
                je_clean = je_clean[valid_month_mask]
                after_month_filter = len(je_clean)
                print(f"过滤无效月份: 过滤前 {before_month_filter} 条，过滤后 {after_month_filter} 条，已过滤 {before_month_filter - after_month_filter} 条")
            except Exception as e:
                print(f"警告: 月份数据清理失败: {e}")
                return pd.DataFrame(), {}
        
        # 处理凭证号格式，提取前缀和数字部分
        je_clean['凭证号_str'] = je_clean['凭证号'].astype(str)
        
        # 提取凭证类型前缀和数字部分
        import re
        def extract_voucher_info(voucher_str):
            # 匹配格式：前缀-数字 或 纯数字
            match = re.match(r'^(.+?)[-_]?(\d+)$', str(voucher_str).strip())
            if match:
                prefix = match.group(1) if match.group(1) != match.group(2) else ''
                number = int(match.group(2))
                return prefix, number
            else:
                # 尝试提取纯数字
                try:
                    return '', int(voucher_str)
                except:
                    return voucher_str, None
        
        # 分批处理以避免内存问题
        voucher_types = []
        voucher_numbers = []
        
        for voucher_str in je_clean['凭证号_str']:
            prefix, number = extract_voucher_info(voucher_str)
            voucher_types.append(prefix)
            voucher_numbers.append(number)
        
        je_clean['凭证类型'] = voucher_types
        je_clean['凭证序号'] = voucher_numbers
        
        # 过滤掉无法解析数字的记录
        before_filter = len(je_clean)
        je_clean = je_clean[je_clean['凭证序号'].notna()]
        after_filter = len(je_clean)
        
        if before_filter > after_filter:
            print(f"过滤无法解析的凭证号: {before_filter - after_filter} 条")
        
        if len(je_clean) == 0:
            print("警告: 没有可解析的凭证号数据")
            return pd.DataFrame(), {}
        
        je_clean['凭证序号'] = je_clean['凭证序号'].astype(int)
        
        # 按年、月、账簿、凭证类型分组检查凭证序号连续性
        gap_results = []
        
        # 构建分组列
        group_cols = ['账簿', '凭证类型']
        if '年' in je_clean.columns:
            group_cols.insert(0, '年')
        if '月' in je_clean.columns:
            group_cols.insert(1, '月')
        
        grouped = je_clean.groupby(group_cols)
        
        for group_key, group in grouped:
            # 获取该组的所有凭证序号并排序
            voucher_numbers = sorted(group['凭证序号'].unique())
            
            if len(voucher_numbers) <= 1:
                continue
                
            # 检查连续性
            gaps = []
            for i in range(len(voucher_numbers) - 1):
                current = voucher_numbers[i]
                next_num = voucher_numbers[i + 1]
                
                # 如果下一个凭证序号不是当前凭证序号+1，则存在跳号
                if next_num != current + 1:
                    gap_start = current + 1
                    gap_end = next_num - 1
                    
                    # 获取分组信息
                    if isinstance(group_key, tuple):
                        if len(group_cols) == 4:  # 年、月、账簿、凭证类型
                            year, month, book, voucher_type = group_key
                        elif len(group_cols) == 3:  # 年、账簿、凭证类型 或 月、账簿、凭证类型
                            if '年' in group_cols:
                                year, book, voucher_type = group_key
                                month = '未知'
                            else:
                                month, book, voucher_type = group_key
                                year = '未知'
                        else:  # 账簿、凭证类型
                            book, voucher_type = group_key
                            year, month = '未知', '未知'
                    else:
                        book = group_key
                        voucher_type = ''
                        year, month = '未知', '未知'
                    
                    # 构造完整的凭证号显示
                    if voucher_type:
                        current_full = f"{voucher_type}-{current}"
                        next_full = f"{voucher_type}-{next_num}"
                        gap_start_full = f"{voucher_type}-{gap_start}"
                        gap_end_full = f"{voucher_type}-{gap_end}"
                    else:
                        current_full = str(current)
                        next_full = str(next_num)
                        gap_start_full = str(gap_start)
                        gap_end_full = str(gap_end)
                    
                    gaps.append({
                        '年': year,
                        '月': month,
                        '账簿': book,
                        '凭证类型': voucher_type,
                        '跳号起始': gap_start_full,
                        '跳号结束': gap_end_full,
                        '跳号数量': gap_end - gap_start + 1,
                        '前一凭证号': current_full,
                        '后一凭证号': next_full
                    })
            
            gap_results.extend(gaps)
        
        gap_df = pd.DataFrame(gap_results)
        
        # 生成统计信息（参考dg_reconciliation_analyzer - 2025.py的实现）
        stats = {
            '检测账簿数': len(je_clean['账簿'].unique()),
            '检测凭证总数': len(je_clean),
            '账簿列表': list(je_clean['账簿'].unique()),
            '凭证类型数': len(je_clean['凭证类型'].unique()) if '凭证类型' in je_clean.columns else 0,
            '凭证类型列表': list(je_clean['凭证类型'].unique()) if '凭证类型' in je_clean.columns else [],
            '检测年月范围': '未知',
            '跳号处数': len(gap_df),
            '总跳号数': gap_df['跳号数量'].sum() if len(gap_df) > 0 else 0
        }
        
        # 安全地计算年月范围（参考dg_reconciliation_analyzer_2023.py的实现）
        if '年' in je_clean.columns and '月' in je_clean.columns and len(je_clean) > 0:
            try:
                # 提取月份的数字部分（处理12A这样的格式）
                def extract_month_number(month_str):
                    import re
                    month_str = str(month_str).strip()
                    # 提取开头的数字部分
                    match = re.match(r'^(\d+)', month_str)
                    if match:
                        return int(match.group(1))
                    return None
                
                je_clean['月_数值'] = je_clean['月'].apply(extract_month_number)
                
                # 过滤掉无效的年月数据
                valid_date_mask = je_clean['年'].notna() & je_clean['月_数值'].notna()
                je_clean_valid_dates = je_clean[valid_date_mask]
                
                if len(je_clean_valid_dates) > 0:
                    min_year = int(je_clean_valid_dates['年'].min())
                    max_year = int(je_clean_valid_dates['年'].max())
                    min_month = int(je_clean_valid_dates['月_数值'].min())
                    max_month = int(je_clean_valid_dates['月_数值'].max())
                    stats['检测年月范围'] = f"{min_year}年{min_month:02d}月 - {max_year}年{max_month:02d}月"
                else:
                    stats['检测年月范围'] = "无有效年月数据"
            except Exception as e:
                print(f"警告: 年月范围计算失败: {e}")
                stats['检测年月范围'] = '年月数据处理异常'
        
        # 生成详细的维度统计信息（按账簿、年、月、凭证类型）
        dimension_stats = []
        grouped = je_clean.groupby(group_cols)
        
        for group_key, group in grouped:
                
            # 获取分组信息
            if isinstance(group_key, tuple):
                if len(group_cols) == 4:  # 年、月、账簿、凭证类型
                    year, month, book, voucher_type = group_key
                elif len(group_cols) == 3:  # 年、账簿、凭证类型 或 月、账簿、凭证类型
                    if '年' in group_cols:
                        year, book, voucher_type = group_key
                        month = '未知'
                    else:
                        month, book, voucher_type = group_key
                        year = '未知'
                else:  # 账簿、凭证类型
                    book, voucher_type = group_key
                    year, month = '未知', '未知'
            else:
                book = group_key
                voucher_type = ''
                year, month = '未知', '未知'
            
            voucher_numbers = sorted(group['凭证序号'].unique())
            if len(voucher_numbers) > 0:
                # 构造完整的凭证号显示
                if voucher_type:
                    min_voucher = f"{voucher_type}-{min(voucher_numbers)}"
                    max_voucher = f"{voucher_type}-{max(voucher_numbers)}"
                else:
                    min_voucher = str(min(voucher_numbers))
                    max_voucher = str(max(voucher_numbers))
                
                # 检查该维度是否有跳号
                has_gap = False
                gap_count = 0
                total_gaps = 0
                
                for i in range(len(voucher_numbers) - 1):
                    current = voucher_numbers[i]
                    next_num = voucher_numbers[i + 1]
                    if next_num != current + 1:
                        has_gap = True
                        gap_count += 1
                        total_gaps += next_num - current - 1
                
                dimension_stats.append({
                    '账簿': book,
                    '年': year,
                    '月': month,
                    '凭证类型': voucher_type,
                    '最小凭证号': min_voucher,
                    '最大凭证号': max_voucher,
                    '凭证数量': len(voucher_numbers),
                    '跳号处数': gap_count,
                    '总跳号数': total_gaps,
                    '检测状态': '❌ 有跳号' if has_gap else '✅ 无跳号'
                })
        
        stats['维度统计'] = dimension_stats
        
        # 更新检测凭证总数为各维度凭证数量的汇总
        total_voucher_count = sum(dim['凭证数量'] for dim in dimension_stats)
        stats['检测凭证总数'] = total_voucher_count
        
        if len(gap_df) > 0:
            print(f"\n=== JE凭证号跳号检测结果 ===")
            print(f"发现跳号情况: {len(gap_df)} 处")
            print(f"总跳号数量: {gap_df['跳号数量'].sum()} 个")
            
            # 按账簿和凭证类型统计跳号情况
            book_type_stats = gap_df.groupby(['账簿', '凭证类型']).agg({
                '跳号数量': ['count', 'sum']
            }).round(2)
            book_type_stats.columns = ['跳号处数', '总跳号数']
            
            print(f"\n=== 各账簿跳号统计 ===")
            for (book, voucher_type), stats_row in book_type_stats.iterrows():
                type_display = f"({voucher_type})" if voucher_type else "(纯数字)"
                print(f"{book} {type_display}: {int(stats_row['跳号处数'])} 处跳号，共 {int(stats_row['总跳号数'])} 个凭证号")
            
            # 总体统计
            book_stats = gap_df.groupby('账簿').agg({
                '跳号数量': ['count', 'sum']
            }).round(2)
            book_stats.columns = ['跳号处数', '总跳号数']
            
            print(f"\n=== 账簿总体统计 ===")
            for book, stats_row in book_stats.iterrows():
                print(f"{book}: {stats_row['跳号处数']} 处跳号，共 {stats_row['总跳号数']} 个凭证号")
            
            # 显示前10个跳号详情
            print(f"\n=== 前10个跳号详情 ===")
            for idx, row in gap_df.head(10).iterrows():
                voucher_type_display = f" - {row['凭证类型']}" if row['凭证类型'] else " - 纯数字"
                print(f"{idx+1}. {row['年']}年{row['月']}月 - {row['账簿']}{voucher_type_display}")
                print(f"   跳号范围: {row['跳号起始']} ~ {row['跳号结束']} (共{row['跳号数量']}个)")
                print(f"   前后凭证号: {row['前一凭证号']} -> {row['后一凭证号']}")
                print()
        else:
            print("\n✅ 未发现凭证号跳号情况，所有凭证号连续")
        
        return gap_df, stats

    def check_voucher_balance(self, target_patterns: List[str]) -> Tuple[pd.DataFrame, Dict]:
        """
        检查所有唯一凭证号的借方发生额与贷方发生额是否一致
        """
        print("正在检查凭证号借贷平衡...")
        
        if self.je_data is None:
            raise ValueError("JE数据未加载")
        
        # 获取配置
        book_col = self.config.get('je_book_column', '账簿')
        voucher_col = self.config.get('je_voucher_column', '凭证号')
        debit_col = self.config.get('je_debit_column', '借方本币')
        credit_col = self.config.get('je_credit_column', '贷方本币')
        year_col = self.config.get('je_year_column', '年')
        month_col = self.config.get('je_month_column', '月')
        
        # 过滤目标账套
        je_filtered = self.je_data.copy()
        if target_patterns:
            pattern_filter = je_filtered[book_col].str.contains('|'.join(target_patterns), na=False)
            je_filtered = je_filtered[pattern_filter]
        
        if je_filtered.empty:
            print("警告: 过滤后的JE数据为空")
            return pd.DataFrame(), {}
        
        # 确保借贷金额为数值类型
        je_filtered[debit_col] = pd.to_numeric(je_filtered[debit_col], errors='coerce').fillna(0)
        je_filtered[credit_col] = pd.to_numeric(je_filtered[credit_col], errors='coerce').fillna(0)
        
        # 按凭证号分组计算借贷合计
        voucher_balance = je_filtered.groupby([book_col, year_col, month_col, voucher_col]).agg({
            debit_col: 'sum',
            credit_col: 'sum'
        }).reset_index()
        
        # 计算借贷差额
        voucher_balance['借贷差额'] = voucher_balance[debit_col] - voucher_balance[credit_col]
        
        # 设置平衡阈值（考虑浮点数精度问题）
        # 使用更严格的阈值，避免将实际平衡的凭证标记为不平衡
        balance_threshold = 1e-2  # 0.01元，更实际的阈值
        
        # 找出不平衡的凭证
        unbalanced_mask = abs(voucher_balance['借贷差额']) > balance_threshold
        unbalanced_vouchers = voucher_balance[unbalanced_mask].copy()
        
        # 调试信息：输出一些可能的精度问题案例
        if len(unbalanced_vouchers) > 0:
            print(f"发现 {len(unbalanced_vouchers)} 个不平衡凭证（阈值: {balance_threshold}）")
            # 显示前几个不平衡凭证的原始差额值
            for i, (idx, row) in enumerate(unbalanced_vouchers.head(3).iterrows()):
                original_diff = voucher_balance.loc[idx, '借贷差额']
                print(f"  凭证 {i+1}: {row[voucher_col]}, 原始借贷差额: {original_diff}, 绝对值: {abs(original_diff)}")
        
        # 重命名列以便输出
        unbalanced_vouchers = unbalanced_vouchers.rename(columns={
            book_col: '账簿',
            year_col: '年',
            month_col: '月',
            voucher_col: '凭证号',
            debit_col: '借方合计',
            credit_col: '贷方合计'
        })
        
        # 格式化金额显示（在过滤之后进行格式化，避免精度问题）
        if len(unbalanced_vouchers) > 0:
            for col in ['借方合计', '贷方合计', '借贷差额']:
                unbalanced_vouchers[col] = unbalanced_vouchers[col].round(2)
            
            # 最终验证：移除格式化后差额为0的记录（可能是由于四舍五入导致的）
            final_unbalanced = unbalanced_vouchers[abs(unbalanced_vouchers['借贷差额']) > 0].copy()
            if len(final_unbalanced) != len(unbalanced_vouchers):
                removed_count = len(unbalanced_vouchers) - len(final_unbalanced)
                print(f"移除了 {removed_count} 个格式化后差额为0的记录")
            unbalanced_vouchers = final_unbalanced
        
        # 统计信息
        total_vouchers = len(voucher_balance)
        unbalanced_count = len(unbalanced_vouchers)
        balanced_count = total_vouchers - unbalanced_count
        
        stats = {
            '总凭证数': total_vouchers,
            '平衡凭证数': balanced_count,
            '不平衡凭证数': unbalanced_count,
            '平衡率': f"{(balanced_count / total_vouchers * 100):.2f}%" if total_vouchers > 0 else "0.00%"
        }
        
        print(f"凭证借贷平衡检查完成: 总凭证数 {total_vouchers}，平衡凭证数 {balanced_count}，不平衡凭证数 {unbalanced_count}")
        
        return unbalanced_vouchers, stats

    def generate_report(self, both_diff: pd.DataFrame, both_no_diff: pd.DataFrame, 
                        je_only: pd.DataFrame, tb_only: pd.DataFrame, 
                        merged_data: pd.DataFrame, gap_df: pd.DataFrame = None, 
                        gap_stats: Dict = None, unbalanced_vouchers: pd.DataFrame = None, 
                        balance_stats: Dict = None) -> str:
        """
        生成对账报告
        """
        output_prefix = self.config.get("output_prefix", "对账报告")
        target_patterns = self.config.get("target_patterns", [])
        pattern_str = "_".join(target_patterns) if target_patterns else "全部账套"
        output_file = f"{output_prefix}_{pattern_str}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        print(f"\n正在生成对账报告: {output_file}")
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 汇总统计
            summary_stats = {
                '差异类型': ['两者都有但存在差异', '两者都有且无差异', '仅在JE中', '仅在TB中', '总计'],
                '记录数': [len(both_diff), len(both_no_diff), len(je_only), len(tb_only), 
                          len(both_diff) + len(both_no_diff) + len(je_only) + len(tb_only)],
                '借方差异金额': [
                    both_diff['借方差异'].sum() if len(both_diff) > 0 else 0,
                    0,  # 无差异记录的差异金额为0
                    je_only['借方本币_JE'].sum() if len(je_only) > 0 else 0,
                    -tb_only['借方本币_TB'].sum() if len(tb_only) > 0 else 0,
                    (both_diff['借方差异'].sum() if len(both_diff) > 0 else 0) + 
                    (je_only['借方本币_JE'].sum() if len(je_only) > 0 else 0) - 
                    (tb_only['借方本币_TB'].sum() if len(tb_only) > 0 else 0)
                ],
                '贷方差异金额': [
                    both_diff['贷方差异'].sum() if len(both_diff) > 0 else 0,
                    0,  # 无差异记录的差异金额为0
                    je_only['贷方本币_JE'].sum() if len(je_only) > 0 else 0,
                    -tb_only['贷方本币_TB'].sum() if len(tb_only) > 0 else 0,
                    (both_diff['贷方差异'].sum() if len(both_diff) > 0 else 0) + 
                    (je_only['贷方本币_JE'].sum() if len(je_only) > 0 else 0) - 
                    (tb_only['贷方本币_TB'].sum() if len(tb_only) > 0 else 0)
                ]
            }
            
            summary_df = pd.DataFrame(summary_stats)
            summary_df.to_excel(writer, sheet_name='汇总统计', index=False)
            
            # 跳号检查统计（参考dg_reconciliation_analyzer - 2025.py的实现）
            if gap_stats:
                # 基本统计信息
                basic_stats = {
                    '统计项目': [
                        '检测账簿数', '检测凭证总数', '凭证类型数', '检测年月范围',
                        '跳号处数', '总跳号数'
                    ],
                    '数值': [
                        gap_stats.get('检测账簿数', 0),
                        gap_stats.get('检测凭证总数', 0),
                        gap_stats.get('凭证类型数', 0),
                        gap_stats.get('检测年月范围', '未知'),
                        gap_stats.get('跳号处数', 0),
                        gap_stats.get('总跳号数', 0)
                    ]
                }
                
                # 账簿列表
                book_list = gap_stats.get('账簿列表', [])
                if book_list:
                    basic_stats['统计项目'].append('检测账簿列表')
                    basic_stats['数值'].append(', '.join(map(str, book_list)))
                
                # 凭证类型列表
                voucher_type_list = gap_stats.get('凭证类型列表', [])
                if voucher_type_list:
                    basic_stats['统计项目'].append('凭证类型列表')
                    basic_stats['数值'].append(', '.join(map(str, voucher_type_list)))
                
                gap_summary_df = pd.DataFrame(basic_stats)
                gap_summary_df.to_excel(writer, sheet_name='跳号检查统计', index=False)
                
                # 维度统计详情
                dimension_stats = gap_stats.get('维度统计', [])
                if dimension_stats:
                    dimension_df = pd.DataFrame(dimension_stats)
                    # 重新排列列的顺序
                    column_order = ['账簿', '年', '月', '凭证类型', '最小凭证号', '最大凭证号', 
                                   '凭证数量', '跳号处数', '总跳号数', '检测状态']
                    dimension_df = dimension_df.reindex(columns=column_order)
                    
                    # 按账簿和月份排序
                    if '账簿' in dimension_df.columns and '月' in dimension_df.columns:
                        # 确保月份列为数值类型以便正确排序
                        dimension_df['月_排序'] = pd.to_numeric(dimension_df['月'], errors='coerce')
                        dimension_df = dimension_df.sort_values(['账簿', '年', '月_排序'], ascending=[True, True, True])
                        # 删除临时排序列
                        dimension_df = dimension_df.drop(columns=['月_排序'])
                    
                    dimension_df.to_excel(writer, sheet_name='跳号统计汇总', index=False)
            
            # 借贷平衡检查统计
            if balance_stats:
                balance_summary = {
                    '统计项目': [
                        '总凭证数', '平衡凭证数', '不平衡凭证数', '平衡率'
                    ],
                    '数值': [
                        balance_stats.get('总凭证数', 0),
                        balance_stats.get('平衡凭证数', 0),
                        balance_stats.get('不平衡凭证数', 0),
                        balance_stats.get('平衡率', '0.00%')
                    ]
                }
                
                balance_summary_df = pd.DataFrame(balance_summary)
                balance_summary_df.to_excel(writer, sheet_name='借贷平衡检查统计', index=False)
                
                # 不平衡凭证明细
                if unbalanced_vouchers is not None and len(unbalanced_vouchers) > 0:
                    unbalanced_vouchers.to_excel(writer, sheet_name='不平衡凭证明细', index=False)
            
            # 排除不需要显示的字段
            exclude_columns = ['je_record_count', 'je_exists', 'tb_record_count', 'tb_exists']
            
            # 各类差异明细
            if len(both_diff) > 0:
                both_diff_display = both_diff.drop(columns=[col for col in exclude_columns if col in both_diff.columns])
                both_diff_display.to_excel(writer, sheet_name='两者都有但存在差异', index=False)
            
            if len(both_no_diff) > 0:
                both_no_diff_display = both_no_diff.drop(columns=[col for col in exclude_columns if col in both_no_diff.columns])
                both_no_diff_display.to_excel(writer, sheet_name='两者都有且无差异', index=False)
            
            if len(je_only) > 0:
                je_only_display = je_only.drop(columns=[col for col in exclude_columns if col in je_only.columns])
                je_only_display.to_excel(writer, sheet_name='仅在JE中', index=False)
            
            if len(tb_only) > 0:
                tb_only_display = tb_only.drop(columns=[col for col in exclude_columns if col in tb_only.columns])
                tb_only_display.to_excel(writer, sheet_name='仅在TB中', index=False)
            
            # 跳号检查详情（参考dg_reconciliation_analyzer - 2025.py的实现）
            if gap_df is not None:
                if len(gap_df) > 0:
                    # 有跳号情况，生成详细报告
                    gap_df.to_excel(writer, sheet_name='跳号检查详情', index=False)
                else:
                    # 无跳号情况，生成检测结果说明
                    no_gap_info = {
                        '检测结果': ['✅ 未发现凭证号跳号情况'],
                        '检测时间': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')],
                        '检测范围': [f"筛选模式: {target_patterns}" if target_patterns else "全部账套"],
                        '统计信息': [
                            f"检测账簿数: {gap_stats.get('检测账簿数', 0)}个, "
                            f"检测凭证总数: {gap_stats.get('检测凭证总数', 0)}个, "
                            f"检测年月范围: {gap_stats.get('检测年月范围', '未知')}, "
                            f"凭证类型数: {gap_stats.get('凭证类型数', 0)}个"
                        ] if gap_stats else ['无统计信息']
                    }
                    
                    # 添加账簿和凭证类型列表
                    if gap_stats:
                        book_list = gap_stats.get('账簿列表', [])
                        voucher_type_list = gap_stats.get('凭证类型列表', [])
                        
                        if book_list:
                            no_gap_info['账簿列表'] = [', '.join(map(str, book_list))]
                        else:
                            no_gap_info['账簿列表'] = ['无']
                            
                        if voucher_type_list:
                            no_gap_info['凭证类型列表'] = [', '.join(map(str, voucher_type_list))]
                        else:
                            no_gap_info['凭证类型列表'] = ['无']
                    
                    no_gap_df = pd.DataFrame(dict([(k, pd.Series(v)) for k, v in no_gap_info.items()]))
                    no_gap_df.to_excel(writer, sheet_name='跳号检查详情', index=False)
            
            # 完整合并数据（排除不需要显示的字段）
            merged_data_display = merged_data.drop(columns=[col for col in exclude_columns if col in merged_data.columns])
            merged_data_display.to_excel(writer, sheet_name='完整对账数据', index=False)
        
        print(f"对账报告已生成: {output_file}")
        return output_file
    
    def run_analysis(self, target_patterns: List[str]) -> str:
        """
        运行完整的对账分析
        """
        print(f"开始对账分析，目标账套模式: {target_patterns}")
        print("=" * 60)
        
        try:
            # 1. 加载数据
            print("\n1. 正在加载数据...")
            self.load_je_files()
            self.load_tb_file()
            
            # 2. 准备数据
            print("\n2. 正在准备数据...")
            je_summary = self.prepare_je_data(target_patterns)
            tb_summary = self.prepare_tb_data(target_patterns)
            
            # 3. 执行对账
            print("\n3. 正在执行对账...")
            both_diff, both_no_diff, je_only, tb_only, merged_data = self.perform_reconciliation(je_summary, tb_summary)
            
            # 4. 执行跳号检查
            print("\n4. 正在执行跳号检查...")
            gap_df, gap_stats = self.check_je_voucher_gaps(target_patterns)
            
            # 5. 执行凭证借贷平衡检查
            print("\n5. 正在执行凭证借贷平衡检查...")
            unbalanced_vouchers, balance_stats = self.check_voucher_balance(target_patterns)
            
            # 6. 生成报告
            print("\n6. 正在生成报告...")
            output_file = self.generate_report(both_diff, both_no_diff, je_only, tb_only, merged_data, gap_df, gap_stats, unbalanced_vouchers, balance_stats)
            
            print(f"\n=== 对账分析完成 ===")
            print(f"报告文件: {output_file}")
            
            return output_file
            
        except Exception as e:
            print(f"\n错误: {e}")
            import traceback
            traceback.print_exc()
            return None

def create_sample_config(config_file: str = "reconciliation_config.json"):
    """
    创建示例配置文件
    """
    sample_config = {
        "target_patterns": ["COMPANY_PATTERN_1", "COMPANY_PATTERN_2"],
        "je_files": ["2025je1-6.xlsx", "2025je7-12.xlsx"],
        "tb_file": "tb2025.xlsx",
        "output_prefix": "对账报告",
        "threshold": 0.01,
        "je_columns": {
            "book": "账簿",
            "subject": "科目",
            "debit": "借方本币",
            "credit": "贷方本币"
        },
        "tb_columns": {
            "book": ["核算账簿名称", "主体账簿", "账簿"],
            "account_code": "科目编码",
            "debit": ["本期借方.1", "本期借方发生.1", "本期借方", "借方累计.1", "借方累计"],
            "credit": ["本期贷方.1", "本期贷方发生.1", "本期贷方", "贷方累计.1", "贷方累计"],
            "debit_col_index": None,
            "credit_col_index": None
        },
        "header_row_index": None,
        "default_book": "默认账簿",
        "summary_patterns": [],
        "filter_invalid_codes": ["总计", "核算账簿累计", "合计", "nan", "币种累计", "科目编码"],
        "filter_patterns": ["币种累计", "核算单位", "制单人", "打印时间"]
    }
    
    with open(config_file, 'w', encoding='utf-8') as f:
        json.dump(sample_config, f, ensure_ascii=False, indent=2)
    
    print(f"示例配置文件已创建: {config_file}")

def main():
    """
    主函数
    """
    parser = argparse.ArgumentParser(description='通用对账分析工具')
    parser.add_argument('--config', '-c', help='配置文件路径')
    parser.add_argument('--patterns', '-p', nargs='+', help='账套筛选模式列表')
    parser.add_argument('--je-files', '-j', nargs='+', help='JE文件路径列表')
    parser.add_argument('--tb-file', '-t', help='TB文件路径')
    parser.add_argument('--create-config', action='store_true', help='创建示例配置文件')
    parser.add_argument('--output-prefix', '-o', help='输出文件前缀')
    
    args = parser.parse_args()
    
    if args.create_config:
        create_sample_config()
        return
    
    # 加载配置
    config = ReconciliationConfig(args.config)
    
    # 命令行参数覆盖配置文件
    if args.patterns:
        config.set_target_patterns(args.patterns)
    
    if args.je_files and args.tb_file:
        config.set_files(args.je_files, args.tb_file)
    
    if args.output_prefix:
        config.config["output_prefix"] = args.output_prefix
    
    # 检查必要参数
    target_patterns = config.get("target_patterns", [])
    je_files = config.get("je_files", [])
    tb_file = config.get("tb_file", "")
    
    if not je_files or not tb_file:
        print("错误: 必须指定JE文件和TB文件")
        print("使用 --help 查看帮助信息")
        print("使用 --create-config 创建示例配置文件")
        return
    
    # 运行分析
    analyzer = GeneralReconciliationAnalyzer(config)
    result = analyzer.run_analysis(target_patterns)
    
    if result:
        print(f"\n分析完成，报告文件: {result}")
    else:
        print("\n分析失败")

if __name__ == "__main__":
    main()