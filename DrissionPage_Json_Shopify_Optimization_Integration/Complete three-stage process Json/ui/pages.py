"""三阶段页面。业务执行由 app 层调度。

注意：本模块已被 app/window.py 中内联的 LinkPage / ProductPage / ProcessingPage
取代，当前无任何引用，保留仅供对照旧版布局。
"""
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog


class Page(ttk.Frame):
    """页面基类：统一维护「运行中需禁用」的控件列表。"""

    def __init__(self, parent, title):
        """渲染页面标题，并初始化待托管控件列表。"""
        super().__init__(parent, padding=24)
        self.controls = []
        ttk.Label(self, text=title, style='Heading.TLabel').pack(anchor='w', pady=(0, 20))

    def path(self, text, variable, command=None):
        """渲染一行「标签 + 路径输入框 + 选择文件夹按钮」。

        选定目录后写入 variable；若给了 command（如读取文件列表）则一并触发。
        """
        row = ttk.Frame(self)
        row.pack(fill='x', pady=(0, 14))
        ttk.Label(row, text=text, width=12).pack(side='left')
        entry = ttk.Entry(row, textvariable=variable)
        entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        def browse():
            """弹出目录选择框，选中后回填变量并可选地触发后续动作。"""
            path = filedialog.askdirectory(parent=self, title=text)
            if path:
                variable.set(path)
                if command:
                    command()
        button = ttk.Button(row, text='选择文件夹', command=browse)
        button.pack(side='left')
        self.controls.extend([entry, button])

    def actions(self, start, handoff=None):
        """渲染底部操作区：主按钮（开始）与可选的「交给下一阶段」按钮。"""
        row = ttk.Frame(self)
        row.pack(fill='x', pady=(16, 0), side='bottom')
        if handoff:
            button = ttk.Button(row, text='交给下一阶段 →', command=handoff)
            button.pack(side='right', padx=(12, 0))
            self.controls.append(button)
        button = ttk.Button(row, text=start[0], command=start[1], style='Primary.TButton')
        button.pack(side='right')
        self.controls.append(button)

    def busy(self, value):
        """运行中禁用（True）或恢复（False）本页登记的全部控件。"""
        for control in self.controls:
            control.configure(state='disabled' if value else 'normal')


class LinkPage(Page):
    """① 链接采集页。"""

    def __init__(self, parent, app):
        """构建保存目录行、页数/XPath 选项与分类列表文本框。"""
        super().__init__(parent, '① 链接采集')
        self.path('链接保存目录', app.vars['link_output'])
        self.actions(('开始链接采集', app.start_links), lambda: app.handoff('links'))
        options = ttk.Frame(self)
        options.pack(fill='x', pady=(0, 14))
        for label, key, width in [('每分类页数', 'max_pages', 8), ('XPath', 'xpath', 36)]:
            ttk.Label(options, text=label).pack(side='left', padx=(0, 8))
            entry = ttk.Entry(options, textvariable=app.vars[key], width=width)
            entry.pack(side='left', padx=(0, 18), fill='x', expand=key=='xpath')
            self.controls.append(entry)
        ttk.Label(self, text='分类列表').pack(anchor='w', pady=(0, 8))
        self.text = scrolledtext.ScrolledText(self, wrap='none', height=16, relief='flat', padx=12, pady=10)
        self.text.pack(fill='both', expand=True)
        self.controls.append(self.text)


class ProductPage(Page):
    """② 商品采集页。"""

    def __init__(self, parent, app):
        """构建链接表目录行、工具条与「文件 / 分类」双层树。"""
        super().__init__(parent, '② 商品采集')
        self.path('链接表目录', app.vars['product_input'], app.scan_files)
        self.actions(('开始商品采集', app.start_products), lambda: app.handoff('products'))
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', pady=(0, 12))
        for text, command in [('读取文件夹', app.scan_files), ('编辑图片规则', app.edit_filter), ('全部展开', lambda: self.expand(True)), ('全部折叠', lambda: self.expand(False))]:
            button = ttk.Button(toolbar, text=text, command=command)
            button.pack(side='left', padx=(0, 8))
            self.controls.append(button)
        self.tree = ttk.Treeview(self, columns=('keep','skip','status'), selectmode='browse')
        self.tree.heading('#0', text='文件 / 分类')
        self.tree.heading('keep', text='保留图片')
        self.tree.heading('skip', text='跳过图片')
        self.tree.heading('status', text='状态')
        self.tree.column('#0', width=420)
        for name in ('keep','skip','status'):
            self.tree.column(name, width=100, stretch=False)
        bar = ttk.Scrollbar(self, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        bar.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)
        self.tree.bind('<Double-1>', lambda event: app.edit_filter())

    def expand(self, value):
        """展开（True）或折叠（False）所有文件节点。"""
        for item in self.tree.get_children():
            self.tree.item(item, open=value)

    def populate(self, files):
        """按读取结果重建文件树，节点 iid 约定为 f{i} / f{i}c{j}。"""
        self.tree.delete(*self.tree.get_children())
        from pathlib import Path
        for i, info in enumerate(files):
            self.tree.insert('', 'end', iid=f'f{i}', text=Path(info['path']).name, open=True,
                             values=(info['keep'], info['skip'], '待运行'))
            for j, (title, rule) in enumerate(info.get('categories', {}).items()):
                self.tree.insert(f'f{i}', 'end', iid=f'f{i}c{j}', text=title or '(无分类)',
                                 values=(rule['keep'], rule['skip'], ''))


class ProcessingPage(Page):
    """③ 数据规范与导出页。"""

    def __init__(self, parent, app):
        """构建中间表目录行、品牌识别展示、字段规则表与结果入口。"""
        super().__init__(parent, '③ 数据规范与导出')
        self.path('中间表目录', app.vars['processing_input'])
        ttk.Label(self, textvariable=app.brand_var, style='Accent.TLabel').pack(anchor='w', pady=(0, 16))
        self.actions(('开始规范并转换', app.start_processing))
        table = ttk.Treeview(self, columns=('rule',), height=7, selectmode='none')
        table.heading('#0', text='字段')
        table.heading('rule', text='当前规则')
        table.column('#0', width=110, stretch=False)
        table.column('rule', width=600)
        for field, rule in [('title','符号清理 · Title Case · 两级分类 · 100 字符'),('name','来源品牌替换 · 符号清理 · Title Case'),('price1','数字类型 · 两位小数'),('price2','两位小数 · 低于 price1 时设为其 1.1 倍'),('src_links','独有图片前四张 + 与 styles1 共有的图片'),('styles1 / 其他','保持原值')]:
            table.insert('', 'end', text=field, values=(rule,))
        table.pack(fill='both', expand=True)
        ttk.Label(self, textvariable=app.result_var, wraplength=720).pack(anchor='w', pady=12)
        ttk.Button(self, text='打开结果目录', command=app.open_results).pack(anchor='e')
