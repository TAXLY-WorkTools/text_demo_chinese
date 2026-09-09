# checker.py
import pandas as pd

def check_cutoff(df):
    """程序2：已提足检查 - 已提足但未提满（补提检查）"""
    results = []
    if df is None or df.empty:
        return pd.DataFrame()

    for idx, row in df.iterrows():
        cost = row.get('original_cost', 0)
        if pd.isna(cost) or cost <= 0:
            continue

        life_years = row.get('useful_life', 0)
        if pd.isna(life_years) or life_years <= 0:
            continue

        rate = row.get('residual_rate', 0)
        if pd.isna(rate):
            rate = 0
        salvage = cost * rate / 100.0
        depreciable_base = cost - salvage

        dep_months = row.get('depreciated_months', 0)
        if pd.isna(dep_months):
            continue

        book_accum = row.get('accumulated_dep', 0)
        if pd.isna(book_accum):
            book_accum = 0

        is_fully = dep_months >= life_years * 12
        if not is_fully:
            continue

        short_amount = round(depreciable_base - book_accum, 2)
        if short_amount > 0.01:
            results.append({
                '资产编号': row.get('asset_id', ''),
                '资产名称': row.get('asset_name', ''),
                '入账日期': row.get('purchase_date', ''),
                '资产原值': round(cost, 2),
                '折旧年限': life_years,
                '已提月份': dep_months,
                '总折旧月数': life_years * 12,
                '净残值': round(salvage, 2),
                '可计提总额': round(depreciable_base, 2),
                '账面累计折旧': round(book_accum, 2),
                '应达到累计折旧': round(depreciable_base, 2),
                '少提金额': short_amount,
                '异常类型': '已提足但未提满（少提折旧）',
                '审计建议': f"补提少计折旧 {short_amount:.2f} 元",
            })

    return pd.DataFrame(results)


def check_over_depreciated(df):
    """程序2辅助（已提足且已提满），返回空DataFrame"""
    return pd.DataFrame()