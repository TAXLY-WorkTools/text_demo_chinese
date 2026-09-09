# calculator.py
import pandas as pd
from data_loader import get_audit_base_date

# ===== 无形资产关键词 =====
INTANGIBLE_KEYWORDS = ['软件', '系统', '程序', 'license', '授权', '数据', '平台', 'SaaS']


def is_intangible(asset_name):
    if not asset_name or not isinstance(asset_name, str):
        return False
    for kw in INTANGIBLE_KEYWORDS:
        if kw in asset_name:
            return True
    return False


def calculate_depreciation(df):
    """
    程序1：折旧测算差异
    重构后逻辑：
    1. 根据入账日期和计提规则，确定“计提起始月”
    2. 计算从计提起始月到基准日的总月数 → 测算累计折旧
    3. 计算从计提起始月到去年12月的总月数 → 相减得到本年应提月数 → 测算本年折旧
    修复：本年折旧不再被“未来剩余额度”截断
    """
    base_date = get_audit_base_date()
    results = []
    
    # ===== 读取方案中的计提规则 =====
    try:
        from scheme_manager import get_current_scheme
        scheme = get_current_scheme()
        if scheme:
            new_asset_rule = scheme.get('global_params', {}).get('new_asset_rule', 'next_month_start')
        else:
            new_asset_rule = 'next_month_start'
    except Exception:
        new_asset_rule = 'next_month_start'
    
    print(f"【计提规则】当前使用的规则: {new_asset_rule}")
    
    for idx, row in df.iterrows():
        cost = row['original_cost']
        life_years = row['useful_life']
        rate = row['residual_rate'] / 100.0 if pd.notna(row['residual_rate']) else 0
        salvage = cost * rate
        depreciable_base = cost - salvage
        monthly_dep_theory = depreciable_base / (life_years * 12) if life_years > 0 else 0
        total_months = life_years * 12

        purchase = row['purchase_date']
        if pd.isna(purchase):
            continue

        asset_name = row.get('asset_name', '')
        is_intang = is_intangible(asset_name)

        # ============================================================
        # 第一步：确定“计提起始月”（根据计提规则）
        # ============================================================
        if new_asset_rule == 'current_month_start':
            start_year = purchase.year
            start_month = purchase.month
        elif new_asset_rule == 'next_month_start':
            start_year = purchase.year
            start_month = purchase.month + 1
            if start_month > 12:
                start_month = 1
                start_year += 1
        elif new_asset_rule == 'by_asset_type':
            if is_intang:
                start_year = purchase.year
                start_month = purchase.month
            else:
                start_year = purchase.year
                start_month = purchase.month + 1
                if start_month > 12:
                    start_month = 1
                    start_year += 1
        else:
            start_year = purchase.year
            start_month = purchase.month + 1
            if start_month > 12:
                start_month = 1
                start_year += 1

        # ============================================================
        # 第二步：计算累计应提月数（从计提起始月到基准日）
        # ============================================================
        end_year = base_date.year
        end_month = base_date.month

        if start_year > end_year or (start_year == end_year and start_month > end_month):
            months = 0
        else:
            months = (end_year - start_year) * 12 + (end_month - start_month) + 1

        if months > total_months:
            months = total_months

        # ============================================================
        # 第三步：计算本年应提月数（修复：正确计算“本年之前已提月数”）
        # ============================================================
        if start_year < base_date.year:
            prev_year = base_date.year - 1
            months_before_year_start = (prev_year - start_year) * 12 + (12 - start_month) + 1
        else:
            months_before_year_start = 0

        months_this_year = months - months_before_year_start

        months_passed_this_year = base_date.month
        if months_this_year > months_passed_this_year:
            months_this_year = months_passed_this_year
        if months_this_year < 0:
            months_this_year = 0

        # ============================================================
        # 第四步：计算测算折旧额（核心修复）
        # ============================================================
        accumulated_dep_theory = min(monthly_dep_theory * months, depreciable_base)
        
        # ===== 修复：本年折旧不应被“未来剩余额度”截断 =====
        # 判断资产是否已提足
        if months >= total_months:
            # 资产已提足（或将在基准日时提足），计算本年实际应提月数
            if months_before_year_start >= total_months:
                # 资产在本年之前已提足，本年应提0
                months_this_year_actual = 0
            else:
                # 本年应提月数 = 总月数 - 本年之前已提月数
                months_this_year_actual = total_months - months_before_year_start
                if months_this_year_actual > months_passed_this_year:
                    months_this_year_actual = months_passed_this_year
                if months_this_year_actual < 0:
                    months_this_year_actual = 0
            dep_this_year_theory = monthly_dep_theory * months_this_year_actual
        else:
            # 资产尚未提足，本年应提就是理论计算值
            dep_this_year_theory = monthly_dep_theory * months_this_year

        # ===== 账面数据 =====
        book_accum = row['accumulated_dep'] if pd.notna(row['accumulated_dep']) else 0
        book_annual = row['annual_dep'] if pd.notna(row['annual_dep']) else 0

        diff_accum = round(book_accum - accumulated_dep_theory, 2)
        diff_annual = round(book_annual - dep_this_year_theory, 2)

        # ===== 已提月份 =====
        dep_months = row.get('depreciated_months', 0)
        if pd.isna(dep_months) or dep_months == 0:
            dep_months = months
        try:
            dep_months = int(dep_months)
        except (ValueError, TypeError):
            dep_months = months

        is_fully_depreciated = dep_months >= life_years * 12

        # ===== 净残值校验 =====
        if is_fully_depreciated:
            expected_accum = depreciable_base
            residual_check_diff = book_accum - expected_accum
            residual_check_status = "已提足校验通过" if abs(residual_check_diff) < 0.01 else "已提足校验异常"
        else:
            expected_accum = None
            residual_check_diff = None
            residual_check_status = "尚未提足"

        asset_type = '无形资产' if is_intang else '固定资产'
        calc_status = '已提足' if is_fully_depreciated else '正常计提中'

        results.append({
            '资产编号': row.get('asset_id', ''),
            '资产名称': asset_name,
            '资产类型': asset_type,
            '入账日期': purchase.strftime('%Y-%m-%d'),
            '资产原值': round(cost, 2),
            '折旧年限': life_years,
            '净残值率(%)': round(rate * 100, 2),
            '账面累计折旧': round(book_accum, 2),
            '测算累计折旧': round(accumulated_dep_theory, 2),
            '累计折旧差异': diff_accum,
            '账面本年折旧': round(book_annual, 2),
            '测算本年折旧': round(dep_this_year_theory, 2),
            '本年折旧差异': diff_annual,
            '理论月折旧': round(monthly_dep_theory, 2),
            '已提月份': dep_months,
            '是否已提足': '是' if is_fully_depreciated else '否',
            '净残值校验状态': residual_check_status,
            '净残值校验差异': round(residual_check_diff, 2) if residual_check_diff is not None else None,
            '测算状态': calc_status,
        })
    
    return pd.DataFrame(results)