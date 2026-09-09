# data_loader.py
import pandas as pd
from pathlib import Path
from config import THRESHOLDS
from field_matcher import load_or_create_mapping, REQUIRED_FIELDS, FIELD_DISPLAY_NAMES


def get_audit_base_date():
    return pd.to_datetime(THRESHOLDS['audit_base_date'])


def load_dataframe(file_path, header_row):
    """
    根据文件格式选择引擎，返回DataFrame
    - .xls: 使用 xlrd
    - .xlsx: 使用 openpyxl
    """
    ext = Path(file_path).suffix.lower()
    header_idx = header_row - 1

    if ext == '.xls':
        try:
            import xlrd
            wb = xlrd.open_workbook(file_path)
            sheet = wb.sheet_by_index(0)
            
            # 读取表头
            headers = []
            for col_idx in range(sheet.ncols):
                val = sheet.cell_value(header_idx, col_idx)
                headers.append(str(val).strip() if val else '')
            
            # 读取数据行
            data = []
            for row_idx in range(header_idx + 1, sheet.nrows):
                row = []
                for col_idx in range(sheet.ncols):
                    val = sheet.cell_value(row_idx, col_idx)
                    row.append(val)
                data.append(row)
            
            df = pd.DataFrame(data, columns=headers)
            return df
        except ImportError:
            raise ImportError("读取 .xls 文件需要 xlrd，请执行: pip install xlrd==1.2.0")
        except Exception as e:
            raise ValueError(f"读取 .xls 文件失败: {e}")

    else:  # .xlsx
        try:
            df = pd.read_excel(file_path, header=header_idx, engine='openpyxl')
            return df
        except ImportError:
            raise ImportError("读取 .xlsx 文件需要 openpyxl，请执行: pip install openpyxl")
        except Exception as e:
            raise ValueError(f"读取 .xlsx 文件失败: {e}")


def load_data_with_mapping(file_path, header_row, mapping):
    """
    使用已确认的映射加载数据并重命名列
    ===== 修改点：增加列名清洗（strip），解决映射值因空格不匹配的问题 =====
    """
    df = load_dataframe(file_path, header_row)
    
    # ===== 修改点1：清洗DataFrame列名（去除首尾空格） =====
    df.columns = [str(col).strip() for col in df.columns]
    
    # 反向映射：Excel列名 -> 标准字段
    # ===== 修改点2：清洗mapping中的列名（去除首尾空格）后再匹配 =====
    reverse_mapping = {}
    for std_field, excel_col in mapping.items():
        if excel_col and excel_col.strip() in df.columns:
            reverse_mapping[excel_col.strip()] = std_field
    
    df = df.rename(columns=reverse_mapping)
    
    # 检查必填字段
    missing = [f for f in REQUIRED_FIELDS if f not in df.columns]
    if missing:
        display_names = [FIELD_DISPLAY_NAMES.get(f, f) for f in missing]
        raise ValueError(f"必填字段缺失，请检查映射配置: {', '.join(display_names)}")
    
    # 日期转换
    if 'purchase_date' in df.columns:
        df['purchase_date'] = pd.to_datetime(df['purchase_date'], errors='coerce')
    
    # 数值转换
    numeric_cols = ['original_cost', 'useful_life', 'residual_rate',
                    'accumulated_dep', 'annual_dep', 'depreciated_months']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
    
    df = df[df['original_cost'].notna() & df['useful_life'].notna()]
    return df