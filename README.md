# Financial Analysis Tools Suite

## 项目概述 (Project Overview)

本项目是一套专业的财务分析工具集，包含会计分录检查器和对账分析器，支持多年度数据分析和多公司账簿处理。所有工具基于Python开发，使用pandas和openpyxl进行Excel数据处理。

This project is a comprehensive suite of financial analysis tools, including accounting voucher analyzers and reconciliation analyzers, supporting multi-year data analysis and multi-company ledger processing.

## 🛠️ 工具列表 (Tools List)

### 1. 会计分录检查器 (Accounting Voucher Analyzers)

#### 主要脚本 (Main Scripts)
- **`accounting_voucher_analyzer_2025.py`** - 最新版本，支持2022-2025年数据
- **`accounting_voucher_analyzer.py`** - 原版本

#### 年度专用版本 (Year-specific Versions)
- `accounting_voucher_analyzer_2022.py` - 2022年专用版本
- `accounting_voucher_analyzer_2023.py` - 2023年专用版本  
- `accounting_voucher_analyzer_2024.py` - 2024年专用版本

#### 功能特性 (Features)
- ✅ **多年度支持**: 支持2022-2025年会计数据分析
- ✅ **节假日识别**: 内置中国法定节假日和调休工作日数据
- ✅ **工作日验证**: 自动检查分录日期是否为工作日
- ✅ **数据完整性检查**: 验证凭证号、日期、制单人、审核人等必填字段
- ✅ **交互式分析**: 支持按年度分析或合并分析
- ✅ **Excel报告生成**: 自动生成详细的分析报告
- ✅ **异常处理**: 完善的错误处理和数据验证机制

### 2. 对账分析器 (Reconciliation Analyzers)

#### 公司专用分析器 (Company-specific Analyzers)
- **`jx_reconciliation_analyzer.py`** - ***REMOVED***
- **`dg_reconciliation_analyzer.py`** - ***REMOVED***
- **`hd_reconciliation_analyzer.py`** - ***REMOVED***

#### 年度版本 (Year-specific Versions)
每个公司都有对应的年度版本：
- `*_reconciliation_analyzer_2022.py`
- `*_reconciliation_analyzer_2023.py`
- `*_reconciliation_analyzer_2024.py`

#### 功能特性 (Features)
- ✅ **JE与TB对账**: 记账凭证(Journal Entry)与试算平衡表(Trial Balance)对账
- ✅ **智能数据解析**: 自动处理货币格式、科目编码提取
- ✅ **多格式支持**: 支持不同的Excel文件格式和列结构
- ✅ **差异分析**: 识别借贷方差异、缺失记录、重复记录
- ✅ **分类报告**: 生成详细的对账差异分类报告
- ✅ **公司定制**: 针对不同公司账簿进行专门优化

## 📋 系统要求 (System Requirements)

### Python版本
- Python 3.7+

### 依赖包 (Dependencies)
```bash
pip install pandas openpyxl numpy xlsxwriter
```

### 文件格式要求 (File Format Requirements)
- 输入文件: Excel格式 (.xlsx)
- 输出文件: Excel格式 (.xlsx)

## 🚀 使用方法 (Usage)

### 会计分录检查器 (Accounting Voucher Analyzer)

#### 基本用法
```bash
# 使用最新版本分析器
python accounting_voucher_analyzer_2025.py

# 指定年份分析
python accounting_voucher_analyzer_2025.py 2024
python accounting_voucher_analyzer_2025.py 2025

# 合并所有年份分析
python accounting_voucher_analyzer_2025.py all
```

#### 输入文件要求
- 文件名包含年份信息 (如: `2024je.xlsx`, `2025je_Q1.xlsx`)
- 必须包含列: `凭证号`, `日期`, `制单人`, `审核人`
- 支持多个工作表的Excel文件

#### 输出文件
- `会计分录检查报告_YYYYMMDD_HHMMSS.xlsx`

#### 批量文件分析器 (Batch Voucher Analyzer)
```bash
# 分析当前目录下所有JE文件（2022-2025年）
python batch_voucher_analyzer.py

# 只分析指定年份的文件
python batch_voucher_analyzer.py --year 2023

# 指定扫描目录
python batch_voucher_analyzer.py --dir "D:\\财务数据"

# 使用自定义文件模式
python batch_voucher_analyzer.py --pattern "*je*.xlsx"
```

### 对账分析器 (Reconciliation Analyzers)

#### ***REMOVED*** (JX)
```bash
# 运行对账分析
python jx_reconciliation_analyzer.py

# 查看帮助
python jx_reconciliation_analyzer.py help
```

**输入文件:**
- `2025je.xlsx` - 记账凭证数据
- `jxtb2025.xlsx` - 试算平衡表数据

**输出文件:**
- `***REMOVED***对账报告_YYYYMMDD_HHMMSS.xlsx`

#### ***REMOVED*** (DG)
```bash
python dg_reconciliation_analyzer.py
```

**输入文件:**
- `2025je.xlsx` - 记账凭证数据
- `tb2025.xlsx` - 试算平衡表数据

**输出文件:**
- `***REMOVED***对账报告_YYYYMMDD_HHMMSS.xlsx`

#### ***REMOVED*** (HD)
```bash
python hd_reconciliation_analyzer.py
```

**输入文件:**
- `2025je.xlsx` - 记账凭证数据
- `hdtb2025.xlsx` - 试算平衡表数据

**输出文件:**
- `***REMOVED***对账报告_YYYYMMDD_HHMMSS.xlsx`

### 通用对账脚本 CLI (General Reconciliation Script)
```bash
python general_reconciliation_script.py \
    --je-file je_data.xlsx \
    --tb-file tb_data.xlsx \
    --target-pattern "COMPANY_PATTERN"
```

可选参数:
- `--config` 配置文件路径 (JSON)
- `--output-prefix` 输出文件前缀
- `--threshold` 对账阈值
- `--output-dir` 报告输出目录
- 多次使用 `--target-pattern` 以筛选多家公司

示例:
```bash
# 单一公司
python general_reconciliation_script.py \
    --je-file je_2025.xlsx \
    --tb-file tb_2025.xlsx \
    --target-pattern "***REMOVED***" \
    --output-prefix "***REMOVED***_对账报告"

# 多家公司
python general_reconciliation_script.py \
    --je-file je_2025.xlsx \
    --tb-file tb_2025.xlsx \
    --target-pattern "***REMOVED***" \
    --target-pattern "***REMOVED***"

# 使用配置文件
python general_reconciliation_script.py \
    --je-file je_2025.xlsx \
    --tb-file tb_2025.xlsx \
    --config company_config.json \
    --threshold 0.01
```

## 📊 报告内容 (Report Contents)

### 会计分录检查报告
- 📈 **统计汇总**: 总分录数、异常分录数、通过率
- 📅 **日期分析**: 工作日验证、节假日检查
- 👥 **人员分析**: 制单人、审核人统计
- 🔍 **异常明细**: 详细的异常记录列表
- 📋 **年度对比**: 多年度数据对比分析

### 对账分析报告
- 📊 **汇总统计**: 总记录数、匹配数、差异数
- ✅ **无差异记录**: 完全匹配的科目明细
- ❌ **存在差异记录**: 借贷方差异明细
- 📝 **仅JE存在**: 只在记账凭证中存在的记录
- 📋 **仅TB存在**: 只在试算平衡表中存在的记录
- 💰 **金额分析**: 借贷方金额差异统计

## 📑 TB格式适配 (TB Format Support)

### 表头行自动检测
- 支持自动检测TB文件中的表头所在行；可通过 `header_row_index` 指定

### 列索引访问与重复列处理
- 当列名重复时可通过索引访问；支持 `debit_col_index` 与 `credit_col_index`

### 默认账簿
- 当TB文件无账簿列时可设置 `default_book` 以自动补充

### 配置示例
```json
{
  "header_row_index": 0,
  "default_book": "默认账簿",
  "tb_columns": {
    "book": ["核算账簿名称", "账簿"],
    "account_code": "科目编码",
    "debit": ["本期借方.1", "借方累计"],
    "credit": ["本期贷方.1", "贷方累计"],
    "debit_col_index": null,
    "credit_col_index": null
  }
}
```

## 📁 项目结构 (Project Structure)

```
testing/
├── README.md                              # 项目说明文档（整合版）
├── requirements.txt                        # Python依赖包
│
├── 会计分录检查器 (Accounting Voucher Analyzers)
│   ├── accounting_voucher_analyzer_2025.py    # 最新版本 (推荐)
│   ├── accounting_voucher_analyzer.py         # 原版本
│   ├── accounting_voucher_analyzer_2022.py    # 2022年版本
│   ├── accounting_voucher_analyzer_2023.py    # 2023年版本
│   └── accounting_voucher_analyzer_2024.py    # 2024年版本
│
├── 对账分析器 (Reconciliation Analyzers)
│   ├── general_reconciliation_script.py       # 通用CLI脚本
│   ├── jx_reconciliation_analyzer.py          # ***REMOVED*** (最新)
│   ├── dg_reconciliation_analyzer.py          # ***REMOVED*** (最新)
│   ├── hd_reconciliation_analyzer.py          # ***REMOVED*** (最新)
│   ├── *_reconciliation_analyzer_2022.py      # 按年度版本
│   ├── *_reconciliation_analyzer_2023.py      # 按年度版本
│   └── *_reconciliation_analyzer_2024.py      # 按年度版本
│
└── 其他文件 (Other Files)
    ├── .gitignore                             # Git忽略文件
    └── 20250801-pz/                          # 数据目录
```

## 🔧 技术特性 (Technical Features)

### 数据处理能力
- **大文件支持**: 支持处理大型Excel文件
- **内存优化**: 优化的数据加载和处理算法
- **格式兼容**: 支持多种Excel格式和列结构
- **编码处理**: 自动处理中文字符编码

### 错误处理
- **异常捕获**: 完善的异常处理机制
- **数据验证**: 多层次的数据完整性验证
- **错误报告**: 详细的错误信息和建议
- **容错机制**: 部分数据错误不影响整体分析

### 性能优化
- **并行处理**: 支持多线程数据处理
- **缓存机制**: 智能的数据缓存策略
- **增量分析**: 支持增量数据分析
- **资源管理**: 自动的内存和文件资源管理

## 📝 更新日志 (Changelog)

### v2025.1 (Latest)
- ✨ 新增2025年数据支持
- 🔧 优化内存使用效率
- 🐛 修复日期解析bug
- 📊 增强报告格式

### v2024.1
- ✨ 新增多公司对账支持
- 🔧 改进科目编码解析
- 📈 优化统计算法

### v2023.1
- ✨ 初始版本发布
- 🎯 基础会计分录检查功能
- 📊 基础对账分析功能

## 🤝 贡献指南 (Contributing)

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开 Pull Request

## 📄 许可证 (License)

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情。

## 📞 支持与联系 (Support & Contact)

如有问题或建议，请通过以下方式联系：

- 📧 Email: [your-email@example.com]
- 🐛 Issues: [GitHub Issues](https://github.com/your-repo/issues)
- 📖 Wiki: [项目Wiki](https://github.com/your-repo/wiki)

## 🙏 致谢 (Acknowledgments)

感谢所有为本项目做出贡献的开发者和用户。

---

**注意**: 使用前请确保已安装所有必要的依赖包，并准备好符合格式要求的输入文件。

**Note**: Please ensure all required dependencies are installed and input files meet the format requirements before use.