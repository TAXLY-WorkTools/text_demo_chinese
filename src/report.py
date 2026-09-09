# report.py
import pandas as pd
from config import THRESHOLDS


def generate_reports(dep_result, checker_result, analysis_result, output_path='折旧测算差异报告.xlsx'):
    """
    生成报告，包含程序1、程序2、程序3的结果
    """
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # ---- Sheet1: 程序1 ----
        if dep_result is not None and not dep_result.empty:
            dep_result.to_excel(writer, sheet_name='程序1_折旧测算差异', index=False)
        else:
            pd.DataFrame({'提示': ['无数据']}).to_excel(writer, sheet_name='程序1_折旧测算差异', index=False)

        # ---- Sheet2: 程序2 ----
        if checker_result is not None and not checker_result.empty:
            checker_result.to_excel(writer, sheet_name='程序2_已提足检查_未提满补提', index=False)
        else:
            pd.DataFrame({'提示': ['无异常']}).to_excel(writer, sheet_name='程序2_已提足检查_未提满补提', index=False)

        # ---- Sheet3: 程序3 异常分析 ----
        if analysis_result is not None and not analysis_result.empty:
            analysis_result.to_excel(writer, sheet_name='程序3_异常分析', index=False)
        else:
            pd.DataFrame({'提示': ['无异常分析结果']}).to_excel(writer, sheet_name='程序3_异常分析', index=False)

        # ---- Sheet4: 汇总统计 ----
        summary = {}
        if dep_result is not None and not dep_result.empty:
            summary['资产总数'] = len(dep_result)
            summary['累计折旧差异总额'] = dep_result['累计折旧差异'].sum()
            summary['本年折旧差异总额'] = dep_result['本年折旧差异'].sum()
            summary['有差异资产数'] = len(dep_result[abs(dep_result['累计折旧差异']) > 0.5])
            summary['重大差异资产数(>500元)'] = len(dep_result[abs(dep_result['累计折旧差异']) > THRESHOLDS.get('materiality_diff', 500)])
        if checker_result is not None and not checker_result.empty:
            summary['已提足未提满资产数'] = len(checker_result)
        else:
            summary['已提足未提满资产数'] = 0
        if analysis_result is not None and not analysis_result.empty:
            summary['异常分析资产数'] = len(analysis_result)
            if '诊断结论' in analysis_result.columns:
                multi = len(analysis_result[analysis_result['诊断结论'].str.contains('多提', na=False)])
                short = len(analysis_result[analysis_result['诊断结论'].str.contains('少提', na=False)])
                summary['多提折旧资产数'] = multi
                summary['少提折旧资产数'] = short
        summary['审计基准日'] = THRESHOLDS.get('audit_base_date', '2025-12-31')
        pd.DataFrame([summary]).to_excel(writer, sheet_name='汇总统计', index=False)