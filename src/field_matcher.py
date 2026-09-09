# field_matcher.py
import json
from pathlib import Path

KEYWORD_RULES = {
    'asset_id': ['编号', '代码', 'id', '卡号', '编码', '序号', '资产编号'],
    'asset_name': ['名称', '资产名称', '品名', '设备名称'],
    'category': ['类别', '分类', '类别名称', '资产类别', '类型'],
    'purchase_date': ['入账日期', '购入日期', '购置日期', '采购日期', '购买日期', '入账日'],
    'original_cost': ['原值', '资产原值', '原价', '金额', '成本', '入账价值', '资产金额'],
    'useful_life': ['年限', '使用年限', '折旧年限', '寿命', '期限'],
    'residual_rate': ['残值率', '净残值率', '净残值%', '残值%', '预计净残值率'],
    'accumulated_dep': ['累计折旧', '累计折旧额', '已提折旧总额', '累计折'],
    'annual_dep': ['本年计提折旧', '本年折旧', '年折旧额', '年折旧', '本期计提折旧'],
    'depreciated_months': ['已提月份', '已计提月数', '计提月份', '已折月份', '已提折旧月数'],
    'location': ['存放地点', '使用地点', '存放位置', '地点', '所在地'],
    'department': ['部门名称', '使用部门', '所属部门', '部门', '责任部门'],
    'net_value': ['净值', '资产净值', '净额'],
    'new_add_dep_flag': ['新增当月计提', '当月计提', '新增计提标记'],
    'enable_date': ['启用日期', '开始使用日期', '启用日'],
    'quantity': ['数量', '资产数量', '台数', '件数'],
}

# ===== 必须映射字段（8个） =====
REQUIRED_FIELDS = [
    'asset_id',
    'asset_name',
    'purchase_date',
    'original_cost',
    'useful_life',
    'residual_rate',
    'accumulated_dep',
    'annual_dep',
]

# ===== 修改点：添加分类字段的显示名称 =====
FIELD_DISPLAY_NAMES = {
    'asset_id': '资产编号',
    'asset_name': '资产名称',
    'category': '资产类别',
    'purchase_date': '入账日期',
    'original_cost': '资产原值',
    'useful_life': '折旧年限(年)',
    'residual_rate': '净残值率(%)',
    'accumulated_dep': '累计折旧(账面)',
    'annual_dep': '本年计提折旧(账面)',
    'depreciated_months': '已提月份',
    'location': '存放地点',
    'department': '部门名称',
    'net_value': '资产净值',
    'new_add_dep_flag': '新增当月计提标记',
    'enable_date': '启用日期',
    'quantity': '资产数量',
    # ===== 新增：资产类型分类相关字段 =====
    'asset_type': '资产类型',
    'tax_life': '税法最低年限',
    'classify_source': '分类来源',
}


def get_all_keywords():
    all_kw = []
    for keywords in KEYWORD_RULES.values():
        all_kw.extend(keywords)
    return all_kw


def smart_match_columns(excel_columns):
    mapping = {}
    matched = set()
    clean_columns = [
        str(col).strip() for col in excel_columns
        if str(col) != 'nan' and col != '' and col is not None
    ]
    for std_field, keywords in KEYWORD_RULES.items():
        for col in clean_columns:
            if col in matched:
                continue
            col_clean = col.strip()
            for kw in keywords:
                if kw in col_clean or col_clean in kw:
                    mapping[std_field] = col
                    matched.add(col)
                    break
    return mapping


def load_or_create_mapping(excel_columns, config_path='mapping_config.json'):
    if Path(config_path).exists():
        with open(config_path, 'r', encoding='utf-8') as f:
            mapping = json.load(f)
        mapping = {k: v for k, v in mapping.items() if v and v != 'null' and v != ''}
        return mapping
    else:
        auto_mapping = smart_match_columns(excel_columns)
        for field in FIELD_DISPLAY_NAMES.keys():
            if field not in auto_mapping:
                auto_mapping[field] = ''
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(auto_mapping, f, ensure_ascii=False, indent=4)
        return None


def save_mapping(mapping, config_path='mapping_config.json'):
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(mapping, f, ensure_ascii=False, indent=4)