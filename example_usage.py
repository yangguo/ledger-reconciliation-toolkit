"""
Example Usage of the General Reconciliation Analyzer
"""

from general_reconciliation_analyzer import GeneralReconciliationAnalyzer, ReconciliationConfig

def example_config_file_reconciliation():
    """
    Example: Reconciliation using configuration file
    """
    print("=== Configuration File Example ===")
    
    # Create analyzer with company-specific configuration
    # The config file should specify target patterns, output prefix, etc.
    config = ReconciliationConfig('company_config.json')
    analyzer = GeneralReconciliationAnalyzer(config)
    
    # Load data (replace with actual file paths)
    # analyzer.load_je_data('je_file.xlsx')
    # analyzer.load_tb_data('tb_file.xlsx')
    
    # Perform reconciliation
    # result = analyzer.reconcile()
    
    # Generate report
    # analyzer.generate_report(output_path='reports')
    
    print("Reconciliation configuration loaded successfully!")
    print(f"Target patterns: {analyzer.config.get('target_patterns')}")
    print(f"Output prefix: {analyzer.config.get('output_prefix')}")
    print(f"Threshold: {analyzer.config.get('threshold')}")
    print("Example completed!")

def example_dict_config_reconciliation():
    """
    Example: Reconciliation using dictionary configuration
    """
    print("\n=== Dictionary Configuration Example ===")
    
    # Company-specific configuration
    company_config = {
        "target_patterns": ["COMPANY_PATTERN_1", "COMPANY_PATTERN_2"],
        "output_prefix": "COMPANY_NAME_对账报告",
        "threshold": 0.01,
        "default_book": "COMPANY_NAME_账簿"
    }
    
    config = ReconciliationConfig()
    if 'target_patterns' in company_config:
        config.set_target_patterns(company_config['target_patterns'])
    if 'output_prefix' in company_config:
        config.config['output_prefix'] = company_config['output_prefix']
    if 'threshold' in company_config:
        config.config['threshold'] = company_config['threshold']
    if 'default_book' in company_config:
        config.config['default_book'] = company_config['default_book']

    analyzer = GeneralReconciliationAnalyzer(config)
    
    # Load data (replace with actual file paths)
    # analyzer.load_je_data('company_je.xlsx')
    # analyzer.load_tb_data('company_tb.xlsx')
    
    print("Reconciliation configuration loaded successfully!")
    print(f"Target patterns: {analyzer.config.get('target_patterns')}")
    print(f"Output prefix: {analyzer.config.get('output_prefix')}")
    print("Example completed!")

def example_general_reconciliation():
    """
    Example: General Reconciliation without specific patterns
    """
    print("\n=== General Reconciliation Example ===")
    
    # Basic configuration
    general_config = {
        "output_prefix": "通用对账报告",
        "threshold": 0.1,
        "filter_invalid_codes": ["总计", "合计", "nan"]
    }
    
    config = ReconciliationConfig()
    if 'output_prefix' in general_config:
        config.config['output_prefix'] = general_config['output_prefix']
    if 'threshold' in general_config:
        config.config['threshold'] = general_config['threshold']
    if 'filter_invalid_codes' in general_config:
        config.config['filter_invalid_codes'] = general_config['filter_invalid_codes']

    analyzer = GeneralReconciliationAnalyzer(config)
    
    print("General Reconciliation configuration loaded successfully!")
    print(f"Output prefix: {analyzer.config.get('output_prefix')}")
    print(f"Threshold: {analyzer.config.get('threshold')}")
    print("Example completed!")

if __name__ == "__main__":
    example_config_file_reconciliation()
    example_dict_config_reconciliation()
    example_general_reconciliation()
