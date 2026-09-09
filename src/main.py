# main.py
# 入口：支持命令行模式和GUI模式
# 启动时自动加载方案

import sys


def apply_scheme_to_config(scheme_data):
    """
    将方案参数应用到 config.py 的全局配置
    """
    try:
        from config import THRESHOLDS
        gp = scheme_data.get('global_params', {})
        
        THRESHOLDS['audit_base_date'] = gp.get('audit_base_date', '2025-12-31')
        THRESHOLDS['residual_rate_min'] = gp.get('residual_rate_min', 0)
        THRESHOLDS['residual_rate_max'] = gp.get('residual_rate_max', 5)
        THRESHOLDS['materiality_diff'] = gp.get('materiality_threshold', 500)
        
        print(f"✅ 方案已应用到配置: {scheme_data.get('scheme_name', '未命名')}")
        print(f"   基准日: {THRESHOLDS['audit_base_date']}")
        print(f"   残值率范围: {THRESHOLDS['residual_rate_min']}% ~ {THRESHOLDS['residual_rate_max']}%")
        print(f"   重要性阈值: {THRESHOLDS['materiality_diff']} 元")
        
        return True
    except Exception as e:
        print(f"⚠️ 应用方案失败: {e}")
        return False


def run_cli(file_path):
    """命令行模式"""
    from data_loader import load_dataframe, load_data_with_mapping
    from calculator import calculate_depreciation
    from report import generate_reports
    from field_matcher import load_or_create_mapping
    from scheme_manager import get_current_scheme
    
    # 尝试加载方案
    scheme = get_current_scheme()
    if scheme:
        apply_scheme_to_config(scheme)
        print(f"使用方案: {scheme.get('scheme_name', '未命名')}")
    else:
        print("使用默认配置")
    
    print(f"正在加载: {file_path}")
    df = load_dataframe(file_path, 3)
    excel_columns = list(df.columns)
    mapping = load_or_create_mapping(excel_columns)
    if mapping is None:
        print("首次运行，已生成 mapping_config.json，请配置后重试")
        return
    df = load_data_with_mapping(file_path, 3, mapping)
    print(f"加载 {len(df)} 条资产")
    print("执行折旧测算...")
    dep_result = calculate_depreciation(df)
    print("生成报告...")
    generate_reports(dep_result, "折旧测算差异报告.xlsx")
    print("完成！报告已生成：折旧测算差异报告.xlsx")


if __name__ == '__main__':
    # ===== 新增：启动时自动加载方案 =====
    try:
        from scheme_manager import get_current_scheme
        scheme = get_current_scheme()
        if scheme:
            apply_scheme_to_config(scheme)
    except Exception as e:
        print(f"⚠️ 启动加载方案失败: {e}")
    
    if len(sys.argv) > 1:
        run_cli(sys.argv[1])
    else:
        from gui import run_gui
        run_gui()