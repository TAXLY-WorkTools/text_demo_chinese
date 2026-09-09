# analyzer.py
# 异常分析模块（程序3）
# 功能：综合分析累计折旧差异和本年折旧差异，诊断差异原因

import pandas as pd


def analyze_differences(dep_result):
    """
    对折旧测算结果进行异常分析
    综合判断累计差异和本年差异，诊断差异原因
    参数：dep_result (DataFrame) - 程序1的输出结果
    返回：分析结果DataFrame（无调整分录列）
    """
    if dep_result is None or dep_result.empty:
        return pd.DataFrame()

    analysis_list = []

    for idx, row in dep_result.iterrows():
        # 提取基础数据
        asset_id = row.get('资产编号', '')
        asset_name = row.get('资产名称', '')
        diff_accum = row.get('累计折旧差异', 0)
        diff_annual = row.get('本年折旧差异', 0)
        monthly_dep = row.get('理论月折旧', 0)
        book_accum = row.get('账面累计折旧', 0)
        theory_accum = row.get('测算累计折旧', 0)
        is_fully = row.get('是否已提足', '') == '是'
        dep_months = row.get('已提月份', 0)

        # 如果两者都为0，跳过
        if abs(diff_accum) < 0.01 and abs(diff_annual) < 0.01:
            continue

        # 计算差异月数（累计差异推算）
        if monthly_dep > 0:
            diff_months = diff_accum / monthly_dep
        else:
            diff_months = 0

        # 以前年度差异 = 累计差异 - 本年差异
        prior_diff = diff_accum - diff_annual

        # ===== 诊断结论构建 =====
        diagnosis_parts = []

        # ---------- 1. 方向诊断 ----------
        if abs(diff_accum) < 0.01 and abs(diff_annual) > 0.01:
            # 特殊情况：累计无差异，但本年有差异
            if diff_annual > 0:
                diagnosis_parts.append("企业本年多提折旧（累计差异被以前年度少提抵消）")
            else:
                diagnosis_parts.append("企业本年少提折旧（累计差异被以前年度多提抵消）")
        elif diff_accum > 0.01:
            diagnosis_parts.append("企业累计多提折旧")
        elif diff_accum < -0.01:
            diagnosis_parts.append("企业累计少提折旧")
        else:
            diagnosis_parts.append("累计无差异")

        # ---------- 2. 整月性诊断 ----------
        if abs(diff_months) > 0.01 and abs(diff_months - round(diff_months)) < 0.01:
            diff_months_int = abs(round(diff_months))
            if diff_months_int > 0:
                diagnosis_parts.append(f"整月数差异（约{diff_months_int}个月）")
            else:
                diagnosis_parts.append("整月数差异（不足1个月）")
        elif abs(diff_months) > 0.01:
            diagnosis_parts.append(
                f"非整月数差异（约{abs(round(diff_months, 2))}个月），可能存在手工调整或四舍五入误差"
            )

        # ---------- 3. 时期分布诊断 ----------
        if abs(diff_accum) > 0.01:
            if abs(prior_diff) < 0.01:
                diagnosis_parts.append("差异全部发生在当年")
            else:
                diagnosis_parts.append(f"差异中{abs(prior_diff):.2f}元发生在以前年度，{abs(diff_annual):.2f}元发生在当年")

        # ---------- 4. 已提足状态诊断 ----------
        if is_fully and diff_accum > 0.01:
            diagnosis_parts.append("账面已提足但程序测算未提足，企业提前提满折旧")
        elif is_fully and diff_accum < -0.01:
            diagnosis_parts.append("账面已提足但程序测算超提，请检查入账日期或年限")
        elif not is_fully and diff_accum > 0.01:
            diagnosis_parts.append("账面未提足但差异将持续存在于未来期间")
        elif not is_fully and diff_accum < -0.01:
            diagnosis_parts.append("账面未提足但程序测算已超提，可能为数据异常")

        # ---------- 5. 新增：累计无差异但本年有差异的特殊说明 ----------
        if abs(diff_accum) < 0.01 and abs(diff_annual) > 0.01:
            if diff_annual > 0:
                diagnosis_parts.append("【需关注】企业当年计提了折旧，但程序认为当年不应计提（如新增资产当年不应计提）")
            else:
                diagnosis_parts.append("【需关注】企业当年未计提折旧，但程序认为当年应计提")

        # 合并诊断结果
        diagnosis_text = "；".join(diagnosis_parts)

        analysis_list.append({
            '资产编号': asset_id,
            '资产名称': asset_name,
            '账面累计折旧': round(book_accum, 2),
            '测算累计折旧': round(theory_accum, 2),
            '累计折旧差异': round(diff_accum, 2),
            '本年折旧差异': round(diff_annual, 2),
            '以前年度差异': round(prior_diff, 2) if abs(prior_diff) > 0.01 else 0,
            '差异月数（按累计差异推算）': round(diff_months, 2) if abs(diff_months) > 0.01 else 0,
            '月折旧额': round(monthly_dep, 2),
            '是否已提足': '是' if is_fully else '否',
            '已提月份': dep_months,
            '诊断结论': diagnosis_text,
        })

    return pd.DataFrame(analysis_list)