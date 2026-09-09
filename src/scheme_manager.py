# scheme_manager.py
# 完整方案管理模块
# 修改：场景名称改为程序1/2/3，删除场景2

import json
import os
import shutil
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


# ============================================================
# 1. 方案数据结构定义
# ============================================================

DEFAULT_SCHEME = {
    "scheme_name": "新方案",
    "scheme_description": "",
    "version": "1.0",
    "created_date": "",
    "modified_date": "",
    
    "global_params": {
        "audit_base_date": "2025-12-31",
        "depreciation_method": "average",
        "new_asset_rule": "next_month_start",
        "currency_precision": 2,
        "materiality_threshold": 500,
        "residual_rate_min": 0,
        "residual_rate_max": 5,
    },
    
    "scenes": {
        "scene1": {
            "enabled": True,
            "name": "程序1: 折旧测算差异",
            "params": {"monthly_dep_rounding": 2}
        },
        "scene3": {
            "enabled": True,
            "name": "程序2: 已提足检查（未提满补提）",
            "params": {}
        },
        "scene4": {
            "enabled": True,
            "name": "程序3: 异常分析（差异原因诊断）",
            "params": {}
        },
        "scene8": {
            "enabled": False,
            "name": "场景8: 抽凭清单（预留）",
            "params": {
                "sample_size": 20,
                "risk_weight_net_value": 40,
                "risk_weight_diff": 30,
                "risk_weight_residual": 15,
                "risk_weight_over_dep": 10,
                "risk_weight_cutoff": 5
            }
        }
    }
}


# ============================================================
# 2. 方案文件读写工具
# ============================================================

def get_schemes_dir():
    schemes_dir = Path("./schemes")
    schemes_dir.mkdir(exist_ok=True)
    return schemes_dir


def list_schemes():
    schemes_dir = get_schemes_dir()
    scheme_files = []
    for f in schemes_dir.glob("*.json"):
        try:
            with open(f, 'r', encoding='utf-8') as file:
                data = json.load(file)
                scheme_files.append({
                    'filename': f.name,
                    'filepath': str(f),
                    'name': data.get('scheme_name', f.stem),
                    'description': data.get('scheme_description', ''),
                    'created_date': data.get('created_date', ''),
                    'modified_date': data.get('modified_date', ''),
                    'data': data
                })
        except Exception:
            continue
    scheme_files.sort(key=lambda x: x['modified_date'], reverse=True)
    return scheme_files


def load_scheme(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_scheme(filepath, scheme_data):
    scheme_data['modified_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if not scheme_data.get('created_date'):
        scheme_data['created_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(scheme_data, f, ensure_ascii=False, indent=4)


def delete_scheme(filepath):
    if Path(filepath).exists():
        Path(filepath).unlink()


def create_default_scheme(name="默认方案"):
    scheme = json.loads(json.dumps(DEFAULT_SCHEME))
    scheme['scheme_name'] = name
    scheme['created_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    scheme['modified_date'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return scheme


def get_current_scheme():
    current_file = Path("./schemes/.current_scheme")
    if current_file.exists():
        try:
            with open(current_file, 'r', encoding='utf-8') as f:
                scheme_path = f.read().strip()
            if Path(scheme_path).exists():
                return load_scheme(scheme_path)
        except Exception:
            pass
    return None


def set_current_scheme(filepath):
    current_file = Path("./schemes/.current_scheme")
    with open(current_file, 'w', encoding='utf-8') as f:
        f.write(str(filepath))


# ============================================================
# 3. 方案管理窗口类
# ============================================================

class SchemeManager:
    def __init__(self, parent, on_apply_callback=None):
        self.parent = parent
        self.on_apply_callback = on_apply_callback
        self.current_scheme_data = None
        self.current_filepath = None
        self.scheme_list = []
        
        self.window = tk.Toplevel(parent)
        self.window.title("📋 方案管理")
        self.window.geometry("1000x650")
        self.window.minsize(900, 550)
        self.window.transient(parent)
        self.window.grab_set()
        
        self._setup_styles()
        self._build_ui()
        self._refresh_scheme_list()
        self.window.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() - 1000) // 2
        y = (self.window.winfo_screenheight() - 650) // 2
        self.window.geometry(f"+{x}+{y}")
    
    def _setup_styles(self):
        style = ttk.Style()
        style.configure("Header.TLabel", font=('Microsoft YaHei', 11, 'bold'))
        style.configure("Section.TLabel", font=('Microsoft YaHei', 10, 'bold'))
    
    def _build_ui(self):
        main_panel = ttk.PanedWindow(self.window, orient=tk.HORIZONTAL)
        main_panel.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        left_frame = ttk.Frame(main_panel, width=280)
        main_panel.add(left_frame, weight=0)
        right_frame = ttk.Frame(main_panel)
        main_panel.add(right_frame, weight=1)
        
        self._build_left_panel(left_frame)
        self._build_right_panel(right_frame)
    
    def _build_left_panel(self, parent):
        ttk.Label(parent, text="📂 方案列表", style="Header.TLabel").pack(anchor=tk.W, pady=(0, 5))
        
        list_container = ttk.Frame(parent)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(list_container)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.scheme_tree = ttk.Treeview(
            list_container,
            columns=('name', 'date'),
            show='tree',
            yscrollcommand=scrollbar.set,
            height=12
        )
        self.scheme_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.scheme_tree.yview)
        
        self.scheme_tree.heading('#0', text='方案名称')
        self.scheme_tree.column('#0', width=140, minwidth=100)
        self.scheme_tree.heading('date', text='修改时间')
        self.scheme_tree.column('date', width=110)
        
        self.scheme_tree.bind('<<TreeviewSelect>>', self._on_scheme_selected)
        
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=5)
        ttk.Button(btn_frame, text="➕ 新建", command=self._new_scheme).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="🗑️ 删除", command=self._delete_scheme).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="📂 导入", command=self._import_scheme).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="💾 导出", command=self._export_scheme).pack(side=tk.LEFT, padx=2)
    
    def _build_right_panel(self, parent):
        canvas = tk.Canvas(parent, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.detail_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.detail_frame, anchor='nw')
        self.detail_frame.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        
        # ---- 基本信息 ----
        ttk.Label(self.detail_frame, text="📋 方案详情", style="Header.TLabel").grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 10))
        
        ttk.Label(self.detail_frame, text="方案名称:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_name = ttk.Entry(self.detail_frame, width=40)
        self.entry_name.grid(row=1, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(self.detail_frame, text="描述:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_desc = ttk.Entry(self.detail_frame, width=40)
        self.entry_desc.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(self.detail_frame, text="版本:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_version = ttk.Entry(self.detail_frame, width=10)
        self.entry_version.grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(self.detail_frame, text="创建时间:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.label_created = ttk.Label(self.detail_frame, text="")
        self.label_created.grid(row=4, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(self.detail_frame, text="修改时间:").grid(row=5, column=0, sticky=tk.W, padx=5, pady=2)
        self.label_modified = ttk.Label(self.detail_frame, text="")
        self.label_modified.grid(row=5, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Separator(self.detail_frame, orient=tk.HORIZONTAL).grid(row=6, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        # ---- 全局参数 ----
        ttk.Label(self.detail_frame, text="🌐 全局参数", style="Section.TLabel").grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        ttk.Label(self.detail_frame, text="审计基准日:").grid(row=8, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_base_date = ttk.Entry(self.detail_frame, width=12)
        self.entry_base_date.grid(row=8, column=1, sticky=tk.W, padx=5, pady=2)
        self.entry_base_date.insert(0, "2025-12-31")
        
        ttk.Label(self.detail_frame, text="折旧方法:").grid(row=9, column=0, sticky=tk.W, padx=5, pady=2)
        self.combo_method = ttk.Combobox(self.detail_frame, values=['平均年限法', '年数总和法', '双倍余额递减法'], width=20)
        self.combo_method.grid(row=9, column=1, sticky=tk.W, padx=5, pady=2)
        self.combo_method.set('平均年限法')
        
        ttk.Label(self.detail_frame, text="新增资产计提规则:").grid(row=10, column=0, sticky=tk.W, padx=5, pady=2)
        self.combo_asset_rule = ttk.Combobox(self.detail_frame, values=['全部次月计提', '全部当月计提', '按资产类型区分'], width=20)
        self.combo_asset_rule.grid(row=10, column=1, sticky=tk.W, padx=5, pady=2)
        self.combo_asset_rule.set('全部次月计提')
        
        ttk.Label(self.detail_frame, text="重要性阈值(元):").grid(row=11, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_materiality = ttk.Entry(self.detail_frame, width=10)
        self.entry_materiality.grid(row=11, column=1, sticky=tk.W, padx=5, pady=2)
        self.entry_materiality.insert(0, "500")
        
        ttk.Separator(self.detail_frame, orient=tk.HORIZONTAL).grid(row=12, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        # ---- 场景配置 ----
        ttk.Label(self.detail_frame, text="⚙️ 场景配置", style="Section.TLabel").grid(row=13, column=0, columnspan=2, sticky=tk.W, pady=(0, 5))
        
        self.scene_vars = {}
        scene_frame = ttk.Frame(self.detail_frame)
        scene_frame.grid(row=14, column=0, columnspan=2, sticky=tk.W, padx=5, pady=2)
        
        scene_names = [
            ('scene1', '程序1: 折旧测算差异'),
            ('scene3', '程序2: 已提足检查（未提满补提）'),
            ('scene4', '程序3: 异常分析（差异原因诊断）'),
            ('scene8', '场景8: 抽凭清单（预留）'),
        ]
        for i, (key, label) in enumerate(scene_names):
            var = tk.BooleanVar(value=True if key == 'scene1' else False)
            self.scene_vars[key] = var
            chk = ttk.Checkbutton(scene_frame, text=label, variable=var)
            chk.grid(row=i//2, column=i%2, sticky=tk.W, padx=10, pady=2)
        
        # 场景参数
        ttk.Label(self.detail_frame, text="残值率范围下限(%):").grid(row=15, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_residual_min = ttk.Entry(self.detail_frame, width=6)
        self.entry_residual_min.grid(row=15, column=1, sticky=tk.W, padx=5, pady=2)
        self.entry_residual_min.insert(0, "0")
        
        ttk.Label(self.detail_frame, text="残值率范围上限(%):").grid(row=16, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_residual_max = ttk.Entry(self.detail_frame, width=6)
        self.entry_residual_max.grid(row=16, column=1, sticky=tk.W, padx=5, pady=2)
        self.entry_residual_max.insert(0, "5")
        
        ttk.Label(self.detail_frame, text="抽凭样本数(场景8):").grid(row=17, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_sample_size = ttk.Entry(self.detail_frame, width=6)
        self.entry_sample_size.grid(row=17, column=1, sticky=tk.W, padx=5, pady=2)
        self.entry_sample_size.insert(0, "20")
        
        ttk.Separator(self.detail_frame, orient=tk.HORIZONTAL).grid(row=18, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        # ---- 底部按钮 ----
        btn_frame = ttk.Frame(self.detail_frame)
        btn_frame.grid(row=19, column=0, columnspan=2, pady=10)
        ttk.Button(btn_frame, text="💾 保存方案", command=self._save_scheme).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✅ 应用方案", command=self._apply_scheme).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 重置为默认值", command=self._reset_to_default).pack(side=tk.LEFT, padx=5)
        
        self.status_label = ttk.Label(self.detail_frame, text="就绪", foreground='gray')
        self.status_label.grid(row=20, column=0, columnspan=2, sticky=tk.W, pady=5)
    
    # ============================================================
    # 3.1 方案列表操作
    # ============================================================
    
    def _refresh_scheme_list(self):
        """刷新方案列表"""
        for item in self.scheme_tree.get_children():
            self.scheme_tree.delete(item)
        
        self.scheme_list = list_schemes()
        for scheme in self.scheme_list:
            self.scheme_tree.insert(
                '',
                'end',
                text=scheme['name'],
                values=(scheme['modified_date'][:10] if scheme['modified_date'] else ''),
                tags=(scheme['filepath'],)
            )
        
        if self.scheme_list:
            self.scheme_tree.selection_set(self.scheme_tree.get_children()[0])
            self._on_scheme_selected(None)
    
    def _on_scheme_selected(self, event):
        selection = self.scheme_tree.selection()
        if not selection:
            return
        item = selection[0]
        filepath = self.scheme_tree.item(item, 'tags')[0]
        try:
            self.current_filepath = filepath
            self.current_scheme_data = load_scheme(filepath)
            self._display_scheme(self.current_scheme_data)
            self.status_label.config(text=f"已加载: {Path(filepath).name}", foreground='green')
        except Exception as e:
            messagebox.showerror("错误", f"加载方案失败: {e}")
    
    def _display_scheme(self, data):
        self.entry_name.delete(0, tk.END)
        self.entry_name.insert(0, data.get('scheme_name', ''))
        self.entry_desc.delete(0, tk.END)
        self.entry_desc.insert(0, data.get('scheme_description', ''))
        self.entry_version.delete(0, tk.END)
        self.entry_version.insert(0, data.get('version', '1.0'))
        self.label_created.config(text=data.get('created_date', ''))
        self.label_modified.config(text=data.get('modified_date', ''))
        
        gp = data.get('global_params', {})
        self.entry_base_date.delete(0, tk.END)
        self.entry_base_date.insert(0, gp.get('audit_base_date', '2025-12-31'))
        method_map = {'average': '平均年限法', 'sum_of_years': '年数总和法', 'double_declining': '双倍余额递减法'}
        self.combo_method.set(method_map.get(gp.get('depreciation_method', 'average'), '平均年限法'))
        rule_map = {'next_month_start': '全部次月计提', 'current_month_start': '全部当月计提', 'by_asset_type': '按资产类型区分'}
        self.combo_asset_rule.set(rule_map.get(gp.get('new_asset_rule', 'next_month_start'), '全部次月计提'))
        self.entry_materiality.delete(0, tk.END)
        self.entry_materiality.insert(0, str(gp.get('materiality_threshold', 500)))
        self.entry_residual_min.delete(0, tk.END)
        self.entry_residual_min.insert(0, str(gp.get('residual_rate_min', 0)))
        self.entry_residual_max.delete(0, tk.END)
        self.entry_residual_max.insert(0, str(gp.get('residual_rate_max', 5)))
        
        scenes = data.get('scenes', {})
        for key, var in self.scene_vars.items():
            var.set(scenes.get(key, {}).get('enabled', True if key == 'scene1' else False))
        scene8 = scenes.get('scene8', {})
        self.entry_sample_size.delete(0, tk.END)
        self.entry_sample_size.insert(0, str(scene8.get('params', {}).get('sample_size', 20)))
    
    # ============================================================
    # 3.2 方案操作（新建、删除、导入、导出）
    # ============================================================
    
    def _new_scheme(self):
        name = f"新方案_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        scheme = create_default_scheme(name)
        filename = f"{name}.json"
        filepath = get_schemes_dir() / filename
        save_scheme(filepath, scheme)
        self._refresh_scheme_list()
        self.status_label.config(text=f"已创建方案: {filename}", foreground='green')
        for item in self.scheme_tree.get_children():
            if self.scheme_tree.item(item, 'tags')[0] == str(filepath):
                self.scheme_tree.selection_set(item)
                self._on_scheme_selected(None)
                break
    
    def _delete_scheme(self):
        selection = self.scheme_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一个方案")
            return
        item = selection[0]
        filepath = self.scheme_tree.item(item, 'tags')[0]
        name = self.scheme_tree.item(item, 'text')
        if not messagebox.askyesno("确认删除", f"确定要删除方案 '{name}' 吗？此操作不可恢复！"):
            return
        delete_scheme(filepath)
        self._refresh_scheme_list()
        self.status_label.config(text=f"已删除方案: {name}", foreground='orange')
        self.current_scheme_data = None
        self.current_filepath = None
        self._clear_detail()
    
    def _clear_detail(self):
        self.entry_name.delete(0, tk.END)
        self.entry_desc.delete(0, tk.END)
        self.entry_version.delete(0, tk.END)
        self.label_created.config(text="")
        self.label_modified.config(text="")
    
    def _import_scheme(self):
        filepath = filedialog.askopenfilename(
            title="导入方案",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")]
        )
        if not filepath:
            return
        try:
            data = load_scheme(filepath)
            if 'scheme_name' not in data or 'global_params' not in data:
                messagebox.showerror("错误", "无效的方案文件：缺少必要字段")
                return
            dest = get_schemes_dir() / Path(filepath).name
            shutil.copy(filepath, dest)
            self._refresh_scheme_list()
            self.status_label.config(text=f"已导入方案: {Path(filepath).name}", foreground='green')
        except Exception as e:
            messagebox.showerror("错误", f"导入方案失败: {e}")
    
    def _export_scheme(self):
        selection = self.scheme_tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择一个方案")
            return
        item = selection[0]
        filepath = self.scheme_tree.item(item, 'tags')[0]
        save_path = filedialog.asksaveasfilename(
            title="导出方案",
            defaultextension=".json",
            filetypes=[("JSON文件", "*.json"), ("所有文件", "*.*")],
            initialfile=Path(filepath).name
        )
        if not save_path:
            return
        try:
            shutil.copy(filepath, save_path)
            self.status_label.config(text=f"已导出方案: {Path(save_path).name}", foreground='green')
        except Exception as e:
            messagebox.showerror("错误", f"导出方案失败: {e}")
    
    # ============================================================
    # 3.3 方案保存与重置
    # ============================================================
    
    def _collect_form_data(self):
        data = {
            'scheme_name': self.entry_name.get().strip() or "未命名方案",
            'scheme_description': self.entry_desc.get().strip(),
            'version': self.entry_version.get().strip() or "1.0",
            'created_date': self.current_scheme_data.get('created_date', datetime.now().strftime('%Y-%m-%d %H:%M:%S')) if self.current_scheme_data else datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'modified_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }
        method_map = {'平均年限法': 'average', '年数总和法': 'sum_of_years', '双倍余额递减法': 'double_declining'}
        rule_map = {'全部次月计提': 'next_month_start', '全部当月计提': 'current_month_start', '按资产类型区分': 'by_asset_type'}
        data['global_params'] = {
            'audit_base_date': self.entry_base_date.get().strip() or '2025-12-31',
            'depreciation_method': method_map.get(self.combo_method.get(), 'average'),
            'new_asset_rule': rule_map.get(self.combo_asset_rule.get(), 'next_month_start'),
            'currency_precision': 2,
            'materiality_threshold': int(self.entry_materiality.get().strip() or 500),
            'residual_rate_min': float(self.entry_residual_min.get().strip() or 0),
            'residual_rate_max': float(self.entry_residual_max.get().strip() or 5),
        }
        scene_names = {
            'scene1': '程序1: 折旧测算差异',
            'scene3': '程序2: 已提足检查（未提满补提）',
            'scene4': '程序3: 异常分析（差异原因诊断）',
            'scene8': '场景8: 抽凭清单（预留）',
        }
        data['scenes'] = {}
        for key, var in self.scene_vars.items():
            data['scenes'][key] = {
                'enabled': var.get(),
                'name': scene_names.get(key, key),
                'params': {}
            }
        data['scenes']['scene8']['params'] = {
            'sample_size': int(self.entry_sample_size.get().strip() or 20),
            'risk_weight_net_value': 40,
            'risk_weight_diff': 30,
            'risk_weight_residual': 15,
            'risk_weight_over_dep': 10,
            'risk_weight_cutoff': 5,
        }
        return data
    
    def _save_scheme(self):
        if not self.current_filepath:
            messagebox.showinfo("提示", "请先新建一个方案或从列表中选择")
            return
        data = self._collect_form_data()
        try:
            save_scheme(self.current_filepath, data)
            self.current_scheme_data = data
            self._refresh_scheme_list()
            self.status_label.config(text=f"✅ 方案已保存: {Path(self.current_filepath).name}", foreground='green')
            messagebox.showinfo("成功", "方案已保存！")
        except Exception as e:
            messagebox.showerror("错误", f"保存方案失败: {e}")
    
    def _reset_to_default(self):
        if not messagebox.askyesno("确认重置", "将重置当前编辑的所有内容为默认值，确定吗？"):
            return
        default = create_default_scheme("临时方案")
        self._display_scheme(default)
        self.status_label.config(text="已重置为默认值，点击'保存方案'以持久化", foreground='orange')
    
    # ============================================================
    # 3.4 方案应用
    # ============================================================
    
    def _apply_scheme(self):
        data = self._collect_form_data()
        if self.current_filepath:
            try:
                save_scheme(self.current_filepath, data)
                self.current_scheme_data = data
                set_current_scheme(self.current_filepath)
                self.status_label.config(text="✅ 方案已应用", foreground='green')
            except Exception as e:
                messagebox.showerror("错误", f"应用方案失败: {e}")
                return
        else:
            if not messagebox.askyesno("提示", "当前方案尚未保存，是否先保存再应用？"):
                return
            name = f"方案_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            filename = f"{name}.json"
            filepath = get_schemes_dir() / filename
            save_scheme(filepath, data)
            self.current_filepath = str(filepath)
            self.current_scheme_data = data
            set_current_scheme(self.current_filepath)
            self._refresh_scheme_list()
            self.status_label.config(text=f"✅ 方案已保存并应用: {filename}", foreground='green')
        if self.on_apply_callback:
            try:
                self.on_apply_callback(data)
                messagebox.showinfo("成功", "方案已应用到主程序！")
            except Exception as e:
                messagebox.showerror("错误", f"应用方案到主程序失败: {e}")
    
    def _on_close(self):
        self.window.destroy()


# ============================================================
# 4. 独立测试函数
# ============================================================

def test_scheme_manager():
    root = tk.Tk()
    root.title("测试 - 方案管理")
    root.geometry("400x200")
    
    def on_apply(data):
        print("方案已应用:")
        print(json.dumps(data, ensure_ascii=False, indent=2))
    
    ttk.Label(root, text="点击下方按钮打开方案管理", font=('Microsoft YaHei', 14)).pack(pady=40)
    ttk.Button(root, text="📋 打开方案管理", command=lambda: SchemeManager(root, on_apply)).pack(pady=10)
    
    root.mainloop()


if __name__ == '__main__':
    test_scheme_manager()