# asset_classifier.py
# 资产类型自动分类模块
# 功能：基于本地关键字库 + 用户自定义 API（兼容 OpenAI 格式），自动识别资产类型和税法最低年限
# 独立于主程序，通过 gui.py 调用

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import re
from pathlib import Path
from datetime import datetime
import threading
import pandas as pd

# ============================================================
# 1. 关键字库管理
# ============================================================

DEFAULT_KEYWORD_LIBRARY = {
    "rules": [
        # 电子设备
        {"keywords": ["电脑", "笔记本", "台式机", "服务器", "工作站", "主机"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["打印机", "复印机", "扫描仪", "传真机", "一体机"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["交换机", "路由器", "防火墙", "网络设备", "AP", "无线"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["空调", "中央空调", "制冷", "新风"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["显示器", "液晶屏", "显示屏", "投影仪", "幕布"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["摄像机", "相机", "摄像头", "监控"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["音响", "功放", "麦克风", "话筒", "调音台"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["平板", "iPad", "PAD"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["硬盘", "存储", "NAS", "磁盘"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        {"keywords": ["电话", "交换机", "程控"], "asset_type": "电子设备", "tax_life": 3, "priority": 10},
        # 办公设备
        {"keywords": ["办公桌", "办公椅", "转椅", "会议桌", "洽谈桌", "工位"], "asset_type": "办公设备", "tax_life": 5, "priority": 8},
        {"keywords": ["文件柜", "档案柜", "保险柜", "储物柜", "柜子", "书架"], "asset_type": "办公设备", "tax_life": 5, "priority": 8},
        {"keywords": ["沙发", "茶几", "茶台", "博古架", "红木"], "asset_type": "办公设备", "tax_life": 5, "priority": 8},
        {"keywords": ["办公设备", "家具"], "asset_type": "办公设备", "tax_life": 5, "priority": 5},
        # 运输工具
        {"keywords": ["奥迪", "别克", "轿车", "车辆", "汽车", "GL8", "A6", "凯越"], "asset_type": "运输工具", "tax_life": 4, "priority": 9},
        {"keywords": ["运输", "货运", "客车"], "asset_type": "运输工具", "tax_life": 4, "priority": 5},
        # 无形资产
        {"keywords": ["软件", "系统", "程序", "license", "授权", "数据", "平台", "SaaS"], "asset_type": "无形资产", "tax_life": 2, "priority": 9},
        # 机器设备
        {"keywords": ["机床", "生产线", "设备", "机械", "加工"], "asset_type": "机器设备", "tax_life": 10, "priority": 5},
    ],
    "version": "1.0",
    "updated_at": ""
}

# 缓存文件路径
CACHE_FILE = Path("./classifier_cache.json")
KEYWORD_FILE = Path("./keyword_library.json")


def load_keyword_library():
    """加载关键字库"""
    if KEYWORD_FILE.exists():
        try:
            with open(KEYWORD_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    lib = DEFAULT_KEYWORD_LIBRARY.copy()
    lib["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return lib


def save_keyword_library(library):
    """保存关键字库"""
    library["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(KEYWORD_FILE, 'w', encoding='utf-8') as f:
        json.dump(library, f, ensure_ascii=False, indent=4)


def load_cache():
    """加载分类缓存"""
    if CACHE_FILE.exists():
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {"cache": {}, "api_calls": 0, "cache_hits": 0}


def save_cache(cache):
    """保存分类缓存"""
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(cache, f, ensure_ascii=False, indent=4)


# ============================================================
# 2. 分类匹配引擎
# ============================================================

def extract_keywords(text):
    """从文本中提取关键词（用于学习机制）"""
    stop_words = ['有限公司', '公司', '集团', '股份', '有限', '责任']
    for sw in stop_words:
        text = text.replace(sw, '')
    matches = re.findall(r'[\u4e00-\u9fff]+', text)
    return [m for m in matches if len(m) >= 2]


def match_asset_type(asset_name, library):
    """匹配资产类型"""
    if not asset_name or not isinstance(asset_name, str):
        return None, None, 'none', None
    
    rules = sorted(library.get('rules', []), key=lambda x: x.get('priority', 0), reverse=True)
    best_match = None
    best_score = 0
    
    for rule in rules:
        keywords = rule.get('keywords', [])
        for kw in keywords:
            if kw in asset_name:
                score = len(kw)
                if score > best_score:
                    best_score = score
                    best_match = rule
    
    if best_match:
        return (
            best_match['asset_type'],
            best_match['tax_life'],
            'keyword_library',
            best_match.get('keywords', [''])[0]
        )
    
    return None, None, 'none', None


def classify_asset(asset_name, library, cache, api_config):
    """分类一条资产（含缓存、关键字库、API三级匹配）"""
    if asset_name in cache.get('cache', {}):
        cached = cache['cache'][asset_name]
        cache['cache_hits'] = cache.get('cache_hits', 0) + 1
        return cached.get('asset_type'), cached.get('tax_life'), 'cache', None, False
    
    atype, tlife, source, kw = match_asset_type(asset_name, library)
    if atype:
        return atype, tlife, source, kw, False
    
    if api_config.get('enabled', False):
        result = call_api(asset_name, api_config)
        if result:
            cache['api_calls'] = cache.get('api_calls', 0) + 1
            return result.get('asset_type'), result.get('tax_life'), 'api', None, True
    
    return None, None, 'none', None, False


# ============================================================
# 3. API 调用（用户自定义 Endpoint，兼容 OpenAI 格式）
# ============================================================

def call_api(asset_name, api_config):
    """
    通用 API 调用（兼容 OpenAI 格式）
    用户自己填写 endpoint、api_key、model
    """
    if not api_config.get('enabled', False):
        return None
    
    endpoint = api_config.get('endpoint', '').strip()
    api_key = api_config.get('api_key', '').strip()
    model = api_config.get('model', '').strip()
    
    if not endpoint or not model:
        return None
    
    prompt = f"""你是一个固定资产审计助手。请根据以下资产名称，判断其资产类型和对应的税法最低折旧年限。

资产名称：{asset_name}

请只输出JSON格式：
{{
    "asset_type": "电子设备/办公设备/运输工具/机器设备/房屋建筑物/无形资产/其他",
    "tax_life_years": 数字,
    "confidence": "高/中/低",
    "reason": "判断依据"
}}

税法最低年限参考：
- 电子设备：3年（电脑、服务器、打印机、空调等）
- 办公设备：5年（办公桌椅、文件柜、保险柜等）
- 运输工具：4年（汽车、车辆等）
- 机器设备：10年（生产设备等）
- 房屋建筑物：20年
- 无形资产：2年（软件、系统等）
- 其他：无法判断时归为其他"""

    try:
        import openai
        client = openai.OpenAI(
            api_key=api_key,
            base_url=endpoint
        )
        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "你是一个专业的固定资产审计助手。"},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1,
            max_tokens=200
        )
        content = response.choices[0].message.content
        return _parse_api_response(content)
    except ImportError:
        return None
    except Exception:
        return None


def _parse_api_response(content):
    """解析API返回的JSON"""
    try:
        data = json.loads(content)
        if 'asset_type' in data:
            return {
                'asset_type': data['asset_type'],
                'tax_life': data.get('tax_life_years', 3),
                'confidence': data.get('confidence', '中'),
                'reason': data.get('reason', '')
            }
    except json.JSONDecodeError:
        import re
        match = re.search(r'\{[^{}]*\}', content)
        if match:
            try:
                data = json.loads(match.group())
                return {
                    'asset_type': data.get('asset_type', '其他'),
                    'tax_life': data.get('tax_life_years', 3),
                    'confidence': data.get('confidence', '中'),
                    'reason': data.get('reason', '')
                }
            except Exception:
                pass
    return None


# ============================================================
# 4. 分类确认窗口
# ============================================================

class ClassifierWindow:
    def __init__(self, parent, df, api_config=None):
        self.parent = parent
        self.original_df = df
        self.api_config = api_config or {'enabled': False}
        self.result = None
        self.cache = load_cache()
        self.library = load_keyword_library()
        self.classification_results = {}
        self.modified_indices = set()
        
        self._classify_all()
        
        self.window = tk.Toplevel(parent)
        self.window.title("资产类型分类确认")
        self.window.geometry("1100x650")
        self.window.minsize(1000, 550)
        self.window.transient(parent)
        self.window.grab_set()
        
        self._build_ui()
        self._refresh_result_table()
        self._update_summary()
        
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() - 1100) // 2
        y = (self.window.winfo_screenheight() - 650) // 2
        self.window.geometry(f"+{x}+{y}")
    
    def _classify_all(self):
        for idx, row in self.original_df.iterrows():
            name_col = None
            for col in ['资产名称', 'asset_name', '名称', 'name']:
                if col in self.original_df.columns:
                    name_col = col
                    break
            
            if name_col is None:
                continue
            
            asset_name = str(row[name_col]) if row[name_col] is not None else ''
            if not asset_name or asset_name == 'nan':
                continue
            
            atype, tlife, source, kw, pending = classify_asset(
                asset_name, self.library, self.cache, self.api_config
            )
            
            self.classification_results[idx] = {
                'asset_name': asset_name,
                'asset_type': atype,
                'tax_life': tlife,
                'source': source,
                'matched_keyword': kw,
                'pending': pending,
                'modified': False
            }
    
    def _build_ui(self):
        top_frame = ttk.Frame(self.window)
        top_frame.pack(fill=tk.X, padx=10, pady=5)
        
        api_status = "已启用" if self.api_config.get('enabled') else "未启用"
        ttk.Label(top_frame, text=f"API状态: {api_status}  |  缓存命中: {self.cache.get('cache_hits', 0)}  |  API调用: {self.cache.get('api_calls', 0)}").pack(side=tk.LEFT)
        
        ttk.Button(top_frame, text="⚙️ API配置", command=self._open_api_config).pack(side=tk.RIGHT, padx=2)
        ttk.Button(top_frame, text="📂 关键字库管理", command=self._open_keyword_manager).pack(side=tk.RIGHT, padx=2)
        
        paned = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        left_frame = ttk.Frame(paned, width=280)
        paned.add(left_frame, weight=0)
        
        ttk.Label(left_frame, text="📂 关键字库摘要", font=('Microsoft YaHei', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        stats_frame = ttk.LabelFrame(left_frame, text="分类统计", padding=10)
        stats_frame.pack(fill=tk.X, pady=5)
        
        self.stat_labels = {}
        stat_types = [
            ('总资产', 'total'),
            ('关键字库命中', 'keyword'),
            ('API建议', 'api'),
            ('待分类', 'pending'),
            ('已确认', 'confirmed')
        ]
        for i, (label, key) in enumerate(stat_types):
            row = ttk.Frame(stats_frame)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=12, anchor='w').pack(side=tk.LEFT)
            self.stat_labels[key] = ttk.Label(row, text='0', foreground='blue')
            self.stat_labels[key].pack(side=tk.RIGHT)
        
        ttk.Label(left_frame, text=f"规则总数: {len(self.library.get('rules', []))} 条", foreground='gray').pack(anchor=tk.W, pady=5)
        ttk.Label(left_frame, text="💡 提示: 双击分类结果可手动修改", foreground='gray', font=('Microsoft YaHei', 9)).pack(anchor=tk.W)
        
        right_frame = ttk.Frame(paned)
        paned.add(right_frame, weight=1)
        
        ttk.Label(right_frame, text="📋 分类结果", font=('Microsoft YaHei', 10, 'bold')).pack(anchor=tk.W, pady=(0, 5))
        
        table_container = ttk.Frame(right_frame)
        table_container.pack(fill=tk.BOTH, expand=True)
        
        columns = ('序号', '资产名称', '建议分类', '税法年限', '来源', '状态')
        self.result_tree = ttk.Treeview(table_container, columns=columns, show='headings', height=18)
        
        col_widths = {'序号': 50, '资产名称': 250, '建议分类': 130, '税法年限': 80, '来源': 100, '状态': 100}
        for col in columns:
            self.result_tree.heading(col, text=col)
            self.result_tree.column(col, width=col_widths.get(col, 100))
        
        scrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.result_tree.bind('<Double-1>', self._on_double_click)
        
        bottom_frame = ttk.Frame(self.window)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(bottom_frame, text='💡 双击"待确认"或"API建议"行可手动修改分类', foreground='gray').pack(side=tk.LEFT)
        
        ttk.Button(bottom_frame, text="✅ 确认全部并继续", command=self._confirm_all).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="⏭️ 跳过分类", command=self._skip_classify).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="❌ 取消加载", command=self._cancel_load).pack(side=tk.RIGHT, padx=5)
    
    def _refresh_result_table(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        for idx, info in self.classification_results.items():
            asset_name = info.get('asset_name', '')[:30]
            atype = info.get('asset_type') or '待分类'
            tlife = info.get('tax_life') or '-'
            source = info.get('source', 'none')
            pending = info.get('pending', False)
            modified = info.get('modified', False)
            
            if pending and source == 'api':
                status = '⚠️ 待确认(API)'
            elif atype == '待分类' or atype is None:
                status = '❌ 待分类'
            elif modified:
                status = '✏️ 已修改'
            else:
                status = '✅ 已确认'
            
            source_labels = {
                'keyword_library': '关键字库',
                'cache': '缓存',
                'api': 'API建议',
                'none': '未匹配'
            }
            source_display = source_labels.get(source, source)
            
            item = self.result_tree.insert('', 'end', values=(
                idx + 1,
                asset_name,
                atype,
                tlife,
                source_display,
                status
            ), tags=(str(idx),))
            
            if pending or atype == '待分类' or atype is None:
                self.result_tree.tag_configure('pending', background='#FFF3E0')
                self.result_tree.item(item, tags=('pending', str(idx)))
            elif source == 'keyword_library':
                self.result_tree.tag_configure('keyword', background='#E8F5E9')
                self.result_tree.item(item, tags=('keyword', str(idx)))
            elif source == 'api':
                self.result_tree.tag_configure('api', background='#E3F2FD')
                self.result_tree.item(item, tags=('api', str(idx)))
    
    def _update_summary(self):
        total = len(self.classification_results)
        keyword_count = sum(1 for v in self.classification_results.values() if v.get('source') == 'keyword_library')
        api_count = sum(1 for v in self.classification_results.values() if v.get('pending'))
        pending_count = sum(1 for v in self.classification_results.values() if v.get('asset_type') is None)
        confirmed_count = total - pending_count
        
        self.stat_labels['total'].config(text=str(total))
        self.stat_labels['keyword'].config(text=str(keyword_count))
        self.stat_labels['api'].config(text=str(api_count))
        self.stat_labels['pending'].config(text=str(pending_count))
        self.stat_labels['confirmed'].config(text=str(confirmed_count))
    
    def _on_double_click(self, event):
        selected = self.result_tree.selection()
        if not selected:
            return
        
        item = selected[0]
        values = self.result_tree.item(item, 'values')
        idx = int(values[0]) - 1
        
        if idx not in self.classification_results:
            return
        
        info = self.classification_results[idx]
        current_type = info.get('asset_type') or '待分类'
        asset_name = info.get('asset_name', '')
        
        popup = tk.Toplevel(self.window)
        popup.title("修改资产类型")
        popup.geometry("400x300")
        popup.transient(self.window)
        popup.grab_set()
        
        ttk.Label(popup, text=f"资产名称: {asset_name}", font=('Microsoft YaHei', 10)).pack(pady=10)
        ttk.Label(popup, text="选择资产类型:").pack(anchor=tk.W, padx=20)
        
        type_var = tk.StringVar(value=current_type)
        types = ['电子设备', '办公设备', '运输工具', '机器设备', '房屋建筑物', '无形资产', '其他']
        type_combo = ttk.Combobox(popup, values=types, textvariable=type_var, width=20, state='readonly')
        type_combo.pack(pady=5)
        
        ttk.Label(popup, text="税法最低年限:").pack(anchor=tk.W, padx=20)
        life_var = tk.StringVar(value=str(info.get('tax_life') or 3))
        life_combo = ttk.Combobox(popup, values=['2', '3', '4', '5', '10', '20'], textvariable=life_var, width=10, state='readonly')
        life_combo.pack(pady=5)
        
        def confirm():
            new_type = type_var.get().strip()
            new_life = int(life_var.get().strip()) if life_var.get().strip().isdigit() else 3
            
            if new_type and new_type != '待分类':
                self.classification_results[idx]['asset_type'] = new_type
                self.classification_results[idx]['tax_life'] = new_life
                self.classification_results[idx]['modified'] = True
                self.classification_results[idx]['source'] = 'user_modified'
                self.classification_results[idx]['pending'] = False
                
                keywords = extract_keywords(asset_name)
                if keywords:
                    add_rule_to_library(keywords, new_type, new_life)
                
                self.cache['cache'][asset_name] = {'asset_type': new_type, 'tax_life': new_life}
                save_cache(self.cache)
                
                self._refresh_result_table()
                self._update_summary()
                popup.destroy()
            else:
                messagebox.showwarning("提示", "请选择有效的资产类型")
        
        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确认", command=confirm).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="取消", command=popup.destroy).pack(side=tk.LEFT, padx=5)
        
        popup.update_idletasks()
        x = (popup.winfo_screenwidth() - 400) // 2
        y = (popup.winfo_screenheight() - 300) // 2
        popup.geometry(f"+{x}+{y}")
    
    def _open_api_config(self):
        ApiConfigWindow(self.window, self.api_config, self._on_api_config_update)
    
    def _on_api_config_update(self, new_config):
        self.api_config = new_config
        for idx, info in self.classification_results.items():
            if info.get('asset_type') is None and self.api_config.get('enabled'):
                atype, tlife, source, kw, pending = classify_asset(
                    info['asset_name'], self.library, self.cache, self.api_config
                )
                if atype:
                    info['asset_type'] = atype
                    info['tax_life'] = tlife
                    info['source'] = source
                    info['pending'] = pending
        
        self._refresh_result_table()
        self._update_summary()
        save_cache(self.cache)
    
    def _open_keyword_manager(self):
        KeywordManagerWindow(self.window, self.library, self._on_keyword_update)
    
    def _on_keyword_update(self, new_library):
        self.library = new_library
        save_keyword_library(new_library)
        self._classify_all()
        self._refresh_result_table()
        self._update_summary()
    
    def _confirm_all(self):
        for idx, info in self.classification_results.items():
            asset_name = info.get('asset_name')
            atype = info.get('asset_type')
            if atype and asset_name:
                self.cache['cache'][asset_name] = {'asset_type': atype, 'tax_life': info.get('tax_life', 3)}
        save_cache(self.cache)
        
        result_df = self.original_df.copy()
        
        type_list = []
        life_list = []
        source_list = []
        
        for idx, row in result_df.iterrows():
            if idx in self.classification_results:
                info = self.classification_results[idx]
                type_list.append(info.get('asset_type') or '待分类')
                life_list.append(info.get('tax_life') or 0)
                source_list.append(info.get('source') or 'none')
            else:
                type_list.append('待分类')
                life_list.append(0)
                source_list.append('none')
        
        result_df['资产类型'] = type_list
        result_df['税法最低年限'] = life_list
        result_df['分类来源'] = source_list
        
        self.result = result_df
        self.window.destroy()
    
    def _skip_classify(self):
        self.result = self.original_df.copy()
        self.window.destroy()
    
    def _cancel_load(self):
        self.result = None
        self.window.destroy()


# ============================================================
# 5. API配置窗口（用户自定义 Endpoint）
# ============================================================

class ApiConfigWindow:
    def __init__(self, parent, current_config, callback):
        self.parent = parent
        self.current_config = current_config.copy()
        self.callback = callback
        
        self.window = tk.Toplevel(parent)
        self.window.title("API 配置")
        self.window.geometry("550x450")
        self.window.transient(parent)
        self.window.grab_set()
        
        self._build_ui()
    
    def _build_ui(self):
        frame = ttk.Frame(self.window, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 启用API
        self.enabled_var = tk.BooleanVar(value=self.current_config.get('enabled', False))
        ttk.Checkbutton(frame, text="启用API智能分类", variable=self.enabled_var).pack(anchor=tk.W, pady=(0, 10))
        
        # ===== API地址（用户自定义） =====
        ttk.Label(frame, text="API地址 (Endpoint):", font=('Microsoft YaHei', 9, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame, text="支持 OpenAI 兼容格式，如：https://api.openai.com/v1 或 http://localhost:1234/v1", 
                  foreground='gray', font=('Microsoft YaHei', 8)).pack(anchor=tk.W)
        self.endpoint_var = tk.StringVar(value=self.current_config.get('endpoint', 'https://api.openai.com/v1'))
        ttk.Entry(frame, textvariable=self.endpoint_var, width=60).pack(anchor=tk.W, pady=(0, 10))
        
        # API Key
        ttk.Label(frame, text="API Key:", font=('Microsoft YaHei', 9, 'bold')).pack(anchor=tk.W)
        self.api_key_var = tk.StringVar(value=self.current_config.get('api_key', ''))
        ttk.Entry(frame, textvariable=self.api_key_var, width=60, show='*').pack(anchor=tk.W, pady=(0, 10))
        
        # 模型名称
        ttk.Label(frame, text="模型名称 (Model):", font=('Microsoft YaHei', 9, 'bold')).pack(anchor=tk.W)
        ttk.Label(frame, text="如：gpt-4o-mini, qwen-turbo, deepseek-chat, 或本地模型名称", 
                  foreground='gray', font=('Microsoft YaHei', 8)).pack(anchor=tk.W)
        self.model_var = tk.StringVar(value=self.current_config.get('model', 'gpt-4o-mini'))
        ttk.Entry(frame, textvariable=self.model_var, width=40).pack(anchor=tk.W, pady=(0, 10))
        
        # 测试连接
        ttk.Button(frame, text="🔗 测试连接", command=self._test_connection).pack(anchor=tk.W, pady=(0, 10))
        
        # 保存
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(btn_frame, text="💾 保存配置", command=self._save).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self.window.destroy).pack(side=tk.RIGHT, padx=5)
    
    def _test_connection(self):
        if not self.enabled_var.get():
            messagebox.showinfo("提示", "请先启用API")
            return
        
        config = {
            'enabled': True,
            'endpoint': self.endpoint_var.get().strip(),
            'api_key': self.api_key_var.get().strip(),
            'model': self.model_var.get().strip()
        }
        
        if not config['endpoint'] or not config['model']:
            messagebox.showwarning("提示", "请填写API地址和模型名称")
            return
        
        result = call_api("联想笔记本电脑", config)
        if result:
            messagebox.showinfo("测试成功", 
                f"测试结果: {result.get('asset_type')}\n"
                f"税法年限: {result.get('tax_life')}年\n"
                f"置信度: {result.get('confidence', '中')}\n"
                f"判断依据: {result.get('reason', '')}"
            )
        else:
            messagebox.showerror("测试失败", 
                "API连接失败，请检查：\n"
                "1. API地址是否正确\n"
                "2. API Key是否有效\n"
                "3. 模型名称是否存在\n"
                "4. 网络是否通畅"
            )
    
    def _save(self):
        self.current_config = {
            'enabled': self.enabled_var.get(),
            'endpoint': self.endpoint_var.get().strip(),
            'api_key': self.api_key_var.get().strip(),
            'model': self.model_var.get().strip()
        }
        self.callback(self.current_config)
        self.window.destroy()


# ============================================================
# 6. 关键字库管理窗口
# ============================================================

class KeywordManagerWindow:
    def __init__(self, parent, library, callback):
        self.parent = parent
        self.library = library
        self.callback = callback
        
        self.window = tk.Toplevel(parent)
        self.window.title("关键字库管理")
        self.window.geometry("700x450")
        self.window.transient(parent)
        self.window.grab_set()
        
        self._build_ui()
        self._refresh_table()
    
    def _build_ui(self):
        frame = ttk.Frame(self.window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        
        table_frame = ttk.Frame(frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ('关键词', '资产类型', '税法年限', '优先级')
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=15)
        self.tree.heading('关键词', text='关键词（多个用/分隔）')
        self.tree.heading('资产类型', text='资产类型')
        self.tree.heading('税法年限', text='税法年限')
        self.tree.heading('优先级', text='优先级')
        self.tree.column('关键词', width=250)
        self.tree.column('资产类型', width=120)
        self.tree.column('税法年限', width=80)
        self.tree.column('优先级', width=80)
        
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="➕ 添加规则", command=self._add_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ 删除选中", command=self._delete_rule).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📥 导入", command=self._import_rules).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📤 导出", command=self._export_rules).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🔄 重置默认", command=self._reset_default).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="✅ 保存并关闭", command=self._save_and_close).pack(side=tk.RIGHT, padx=2)
    
    def _refresh_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for rule in self.library.get('rules', []):
            keywords = '/'.join(rule.get('keywords', []))
            self.tree.insert('', 'end', values=(
                keywords,
                rule.get('asset_type', ''),
                rule.get('tax_life', ''),
                rule.get('priority', 5)
            ))
    
    def _add_rule(self):
        popup = tk.Toplevel(self.window)
        popup.title("添加规则")
        popup.geometry("400x250")
        popup.transient(self.window)
        popup.grab_set()
        
        ttk.Label(popup, text="关键词（多个用空格分隔）:").pack(anchor=tk.W, padx=20, pady=(10, 2))
        kw_entry = ttk.Entry(popup, width=40)
        kw_entry.pack(padx=20, pady=5)
        
        ttk.Label(popup, text="资产类型:").pack(anchor=tk.W, padx=20, pady=2)
        type_var = tk.StringVar(value='电子设备')
        type_combo = ttk.Combobox(popup, values=['电子设备', '办公设备', '运输工具', '机器设备', '房屋建筑物', '无形资产', '其他'], 
                                  textvariable=type_var, width=20)
        type_combo.pack(padx=20, pady=5)
        
        ttk.Label(popup, text="税法最低年限:").pack(anchor=tk.W, padx=20, pady=2)
        life_var = tk.StringVar(value='3')
        life_combo = ttk.Combobox(popup, values=['2', '3', '4', '5', '10', '20'], textvariable=life_var, width=10)
        life_combo.pack(padx=20, pady=5)
        
        def confirm():
            keywords = [kw.strip() for kw in kw_entry.get().split() if kw.strip()]
            if not keywords:
                messagebox.showwarning("提示", "请输入至少一个关键词")
                return
            atype = type_var.get()
            tlife = int(life_var.get()) if life_var.get().isdigit() else 3
            
            self.library['rules'].append({
                'keywords': keywords,
                'asset_type': atype,
                'tax_life': tlife,
                'priority': 5
            })
            self._refresh_table()
            popup.destroy()
        
        ttk.Button(popup, text="确认添加", command=confirm).pack(pady=10)
    
    def _delete_rule(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("提示", "请先选择要删除的规则")
            return
        
        if not messagebox.askyesno("确认删除", "确定要删除选中的规则吗？"):
            return
        
        for item in selected:
            values = self.tree.item(item, 'values')
            keywords = values[0].split('/')
            
            self.library['rules'] = [
                r for r in self.library.get('rules', [])
                if r.get('keywords') != keywords
            ]
        
        self._refresh_table()
    
    def _import_rules(self):
        filepath = filedialog.askopenfilename(
            title="导入规则",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if not filepath:
            return
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if 'rules' in data:
                self.library['rules'] = data['rules']
                self._refresh_table()
                messagebox.showinfo("成功", f"已导入 {len(data['rules'])} 条规则")
            else:
                messagebox.showerror("错误", "无效的规则文件格式")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {e}")
    
    def _export_rules(self):
        filepath = filedialog.asksaveasfilename(
            title="导出规则",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if not filepath:
            return
        
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump({'rules': self.library.get('rules', [])}, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", "规则已导出")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {e}")
    
    def _reset_default(self):
        if not messagebox.askyesno("确认重置", "将重置为默认规则，当前规则将被覆盖。确定吗？"):
            return
        self.library['rules'] = DEFAULT_KEYWORD_LIBRARY['rules']
        self._refresh_table()
    
    def _save_and_close(self):
        self.callback(self.library)
        self.window.destroy()


# ============================================================
# 7. 辅助函数
# ============================================================

def add_rule_to_library(keywords, asset_type, tax_life):
    """添加规则到关键字库（学习机制）"""
    library = load_keyword_library()
    
    for rule in library.get('rules', []):
        if set(rule.get('keywords', [])) == set(keywords):
            return
    
    library['rules'].append({
        'keywords': keywords,
        'asset_type': asset_type,
        'tax_life': tax_life,
        'priority': 5
    })
    save_keyword_library(library)


# ============================================================
# 8. 主入口函数
# ============================================================

def show_classifier(parent, df, api_config=None):
    """
    显示分类确认窗口
    返回: DataFrame (含分类列) 或 None (用户取消)
    """
    if df is None or df.empty:
        return None
    
    name_cols = ['资产名称', 'asset_name', '名称', 'name']
    has_name = any(col in df.columns for col in name_cols)
    if not has_name:
        messagebox.showwarning("提示", "数据中未找到资产名称列，无法自动分类")
        return None
    
    window = ClassifierWindow(parent, df, api_config)
    parent.wait_window(window.window)
    
    return window.result


def test_classifier():
    """独立测试函数"""
    root = tk.Tk()
    root.title("测试 - 资产分类")
    root.geometry("400x200")
    
    test_df = pd.DataFrame({
        '资产编号': ['001', '002', '003', '004', '005'],
        '资产名称': ['联想笔记本电脑', '奥迪A6', '广联达BIM软件', '1.8米办公桌', '设备A'],
        '资产原值': [5000, 200000, 3000, 1200, 10000],
    })
    
    api_config = {'enabled': False}
    
    def on_classify():
        result = show_classifier(root, test_df, api_config)
        if result is not None:
            print("分类结果:")
            print(result[['资产名称', '资产类型', '税法最低年限']])
    
    ttk.Label(root, text="点击下方按钮测试资产分类", font=('Microsoft YaHei', 14)).pack(pady=40)
    ttk.Button(root, text="🔍 资产分类", command=on_classify).pack(pady=10)
    
    root.mainloop()


if __name__ == '__main__':
    test_classifier()