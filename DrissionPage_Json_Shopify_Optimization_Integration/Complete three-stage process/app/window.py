"""现代化商品采集与数据处理工作台。

说明：
- 保留现有 TaskController / operations / collector / field_rules 等后端调用。
- UI 层重构为 Sidebar + Topbar + Content + Status Bar。
- 不引入第三方 UI 库，仅使用 tkinter / ttk。
"""

from __future__ import annotations

import copy
import os
from pathlib import Path
import queue
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from app.task_paths import INPUT_ROOT, directories, task_path, clean_task_name
from app.task_controller import TaskController, checkpoint
from app.operations import collect_links, collect_products, process_data, convert_data
from stages.link_collection.collector import parse_categories
from stages.product_collection.file_utils import read_categories_from_excel
from stages.data_processing.field_rules.name import extract_brand


APP_DIR = Path(__file__).resolve().parents[1]


COLORS = {
    "bg": "#F6F7FB",
    "surface": "#FFFFFF",
    "surface_soft": "#F8FAFC",
    "sidebar": "#111827",
    "sidebar_hover": "#1F2937",
    "sidebar_active": "#2563EB",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_light": "#EFF6FF",
    "text": "#0F172A",
    "text_secondary": "#64748B",
    "text_muted": "#94A3B8",
    "border": "#E5E7EB",
    "border_strong": "#CBD5E1",
    "success": "#16A34A",
    "success_bg": "#F0FDF4",
    "warning": "#D97706",
    "warning_bg": "#FFFBEB",
    "danger": "#DC2626",
    "danger_bg": "#FEF2F2",
    "console": "#0F172A",
    "console_text": "#CBD5E1",
}


PAGE_META = {
    "links": ("链接采集", "从分类页面批量获取商品链接，并生成标准化链接表。"),
    "products": ("商品采集", "读取链接表并批量采集商品详情与图片数据。"),
    "processing": ("规范与合并", "规范中间表字段并生成合并表。"),
    "conversion": ("转换与导出", "从合并表生成 Shopify / WooCommerce 文件。"),
    "logs": ("运行日志", "实时查看各阶段的执行记录与异常信息。"),
}


def choose_font(root: tk.Misc) -> str:
    """尽量选择平台原生中文 UI 字体。"""
    families = set(root.tk.call("font", "families"))
    candidates = []
    if sys.platform == "darwin":
        candidates = ["PingFang SC", "Helvetica Neue", "Arial"]
    elif os.name == "nt":
        candidates = ["Microsoft YaHei UI", "Microsoft YaHei", "Segoe UI"]
    else:
        candidates = ["Noto Sans CJK SC", "DejaVu Sans", "Arial"]

    for name in candidates:
        if name in families:
            return name
    return "TkDefaultFont"


class BasePage(ttk.Frame):
    """页面公共能力：忙闲状态批量切换、卡片容器与区块标题。"""

    def __init__(self, parent, app: "Workbench"):
        """记录宿主 Workbench，并初始化待托管的“运行中需禁用”控件列表。"""
        super().__init__(parent, style="Page.TFrame")
        self.app = app
        self._busy_widgets: list[tk.Widget] = []

    def track_busy(self, *widgets):
        """登记一批控件，使其在任务运行期间由 busy(True) 统一禁用。"""
        self._busy_widgets.extend(widgets)
        return widgets[-1] if widgets else None

    def busy(self, state: bool):
        """按运行状态启用/禁用所有已登记控件。

        Combobox 需要还原成 readonly 而非 normal，否则会变成可编辑下拉框。
        """
        for widget in self._busy_widgets:
            try:
                if isinstance(widget, ttk.Combobox):
                    widget.configure(state="disabled" if state else ("normal" if getattr(widget, "_editable", False) else "readonly"))
                elif isinstance(widget, tk.Text):
                    widget.configure(state="disabled" if state else "normal")
                else:
                    widget.configure(state="disabled" if state else "normal")
            except tk.TclError:
                pass

    def card(self, parent, *, padding=20):
        """创建带 1px 描边的白色卡片，返回 (外层 Frame, 内层内容 Frame)。"""
        outer = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            bd=0,
        )
        inner = ttk.Frame(outer, style="Card.TFrame", padding=padding)
        inner.pack(fill="both", expand=True)
        return outer, inner

    def section_title(self, parent, title: str, subtitle: str | None = None):
        """在卡片内渲染主标题，可选带一行灰色副标题说明。"""
        ttk.Label(parent, text=title, style="CardTitle.TLabel").pack(anchor="w")
        if subtitle:
            ttk.Label(
                parent,
                text=subtitle,
                style="Muted.TLabel",
                wraplength=760,
            ).pack(anchor="w", pady=(5, 0))


class LinkPage(BasePage):
    """① 链接采集页：分类 URL 输入 + 输出目录/页数/XPath 设置。"""

    def __init__(self, parent, app: "Workbench"):
        """构建分类文本框、采集设置区与「开始采集 / 继续到商品采集」按钮。"""
        super().__init__(parent, app)

        shell = ttk.Frame(self, style="Page.TFrame", padding=(20, 12, 20, 12))
        shell.pack(fill="both", expand=True)
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(0, weight=1)

        # 分类 URL
        card1, body1 = self.card(shell, padding=12)
        card1.grid(row=0, column=0, sticky="nsew")
        self.section_title(body1, "分类 URL")

        toolbar = ttk.Frame(body1, style="Card.TFrame")
        toolbar.pack(fill="x", pady=(4, 0))
        parse_button = ttk.Button(toolbar, text="解析为表格", style="Secondary.TButton", command=self.parse_table)
        parse_button.pack(side="left")
        edit_button = ttk.Button(toolbar, text="返回文本编辑", style="Secondary.TButton", command=self.edit_text)
        edit_button.pack(side="left", padx=8)
        self.track_busy(parse_button, edit_button)
        self.table_mode = False
        self.parsed_categories = []
        self.input_area = ttk.Frame(body1, style="Card.TFrame")
        self.input_area.pack(fill="both", expand=True, pady=(6, 0))
        text_border = tk.Frame(
            self.input_area,
            bg=COLORS["border_strong"],
            padx=1,
            pady=1,
        )
        text_border.pack(fill="both", expand=True)
        self.text_container = text_border

        self.text = tk.Text(
            text_border,
            height=4,
            wrap="word",
            relief="flat",
            bd=0,
            padx=12,
            pady=10,
            bg=COLORS["surface_soft"],
            fg=COLORS["text"],
            insertbackground=COLORS["primary"],
            font=("TkFixedFont", 10),
            undo=True,
        )
        self.text.pack(fill="both", expand=True)
        self.track_busy(self.text)
        self.table_container = ttk.Frame(self.input_area, style="Card.TFrame")
        self.category_table = ttk.Treeview(self.table_container, columns=("title", "url"), show="headings", height=4)
        self.category_table.heading("title", text="分类名称")
        self.category_table.heading("url", text="分类 URL")
        self.category_table.column("title", width=220)
        self.category_table.column("url", width=600)
        bar = ttk.Scrollbar(self.table_container, orient="vertical", command=self.category_table.yview)
        self.category_table.configure(yscrollcommand=bar.set)
        bar.pack(side="right", fill="y")
        self.category_table.pack(fill="both", expand=True)

        # 采集设置
        card2, body2 = self.card(shell, padding=12)
        card2.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        self.section_title(body2, "采集设置")

        settings = ttk.Frame(body2, style="Card.TFrame")
        settings.pack(fill="x", pady=(8, 0))
        settings.columnconfigure(0, weight=1)
        settings.columnconfigure(1, weight=1)
        settings.columnconfigure(2, weight=0)

        # 固定根目录：批次 → 负责人 → 本次任务。
        group_output = ttk.Frame(settings, style="Card.TFrame")
        group_output.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.batch_var = tk.StringVar()
        self.owner_var = tk.StringVar()
        self.task_folder_var = tk.StringVar()
        self.path_preview = tk.StringVar(value="01INPUT_XLSX/")
        self.path_ready = tk.StringVar()
        self.app.vars["link_output"].set("")
        row = ttk.Frame(group_output, style="Card.TFrame")
        row.pack(fill="x")
        row.columnconfigure(0, weight=1)
        row.columnconfigure(1, weight=1)
        for column, (label, variable) in enumerate((("批次", self.batch_var), ("负责人", self.owner_var))):
            ttk.Label(row, text=label, style="FieldLabel.TLabel").grid(row=0, column=column, sticky="w")
            combo = ttk.Combobox(row, textvariable=variable, state="normal", width=14, style="Modern.TCombobox")
            combo.grid(row=1, column=column, sticky="ew", padx=(0, 8), pady=(6, 0))
            combo._editable = True
            self.track_busy(combo)
            if column == 0:
                self.batch_combo = combo
            else:
                self.owner_combo = combo
        refresh = ttk.Button(row, text="刷新", command=self.refresh_batches, style="Secondary.TButton")
        refresh.grid(row=1, column=2)
        ttk.Label(group_output, text="任务文件夹", style="FieldLabel.TLabel").pack(anchor="w", pady=(8, 0))
        task_row = ttk.Frame(group_output, style="Card.TFrame")
        task_row.pack(fill="x")
        entry = ttk.Entry(task_row, textvariable=self.task_folder_var, style="Modern.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        create = ttk.Button(task_row, text="创建文件夹", command=self.create_task_folder, style="Secondary.TButton")
        create.pack(side="left", padx=8)
        ttk.Label(group_output, textvariable=self.path_preview, style="Muted.TLabel", wraplength=480).pack(anchor="w", pady=(8, 0))
        ttk.Label(group_output, textvariable=self.path_ready, style="Muted.TLabel").pack(anchor="w")
        self.track_busy(entry, refresh, create)
        self.batch_var.trace_add("write", lambda *args: self.refresh_owners())
        self.owner_var.trace_add("write", self.update_task_path)
        self.task_folder_var.trace_add("write", self.update_task_path)
        self.refresh_batches()

        # 最大页数
        group_pages = ttk.Frame(settings, style="Card.TFrame")
        group_pages.grid(row=1, column=0, sticky="new", padx=(0, 12), pady=(12, 0))
        ttk.Label(group_pages, text="最大页数（留空不限）", style="FieldLabel.TLabel").pack(anchor="w")
        entry_pages = ttk.Entry(
            group_pages,
            width=12,
            textvariable=app.vars["max_pages"],
            style="Modern.TEntry",
        )
        entry_pages.pack(anchor="w", pady=(4, 0))
        self.track_busy(entry_pages)

        # XPath
        group_xpath = ttk.Frame(settings, style="Card.TFrame")
        group_xpath.grid(row=1, column=1, columnspan=2, sticky="new", pady=(12, 0))
        ttk.Label(group_xpath, text="XPath", style="FieldLabel.TLabel").pack(anchor="w")
        entry_xpath = ttk.Entry(
            group_xpath,
            textvariable=app.vars["xpath"],
            style="Modern.TEntry",
        )
        entry_xpath.pack(fill="x", pady=(4, 0))
        self.track_busy(entry_xpath)

        # 操作区
        actions = ttk.Frame(shell, style="Page.TFrame")
        actions.grid(row=2, column=0, sticky="ew", pady=(10, 0))

        self.next_btn = ttk.Button(
            actions,
            text="继续到商品采集  →",
            style="Secondary.TButton",
            command=lambda: app.handoff("links"),
        )
        self.next_btn.pack(side="right", padx=(8, 0))

        self.start_btn = ttk.Button(
            actions,
            text="开始采集  →",
            style="Primary.TButton",
            command=app.start_links,
        )
        self.start_btn.pack(side="right")
        self.track_busy(self.start_btn, self.next_btn)

    def parse_table(self):
        if self.app.busy:
            return
        try:
            items = parse_categories(self.text.get("1.0", "end-1c"))
            if not items:
                raise ValueError("请输入有效分类 URL")
        except ValueError as exc:
            messagebox.showerror("解析失败", str(exc), parent=self.app.root)
            return
        self.parsed_categories = items
        self.category_table.delete(*self.category_table.get_children())
        for item in items:
            self.category_table.insert("", "end", values=(item.title, item.url))
        self.text_container.pack_forget()
        self.table_container.pack(fill="both", expand=True)
        self.table_mode = True

    def edit_text(self):
        if self.app.busy:
            return
        self.table_container.pack_forget()
        self.text_container.pack(fill="both", expand=True)
        self.table_mode = False

    def category_input(self):
        if self.table_mode:
            # 使用字面量保留名称内的逗号、引号等字符。
            return repr([(item.title, item.url) for item in self.parsed_categories])
        return self.text.get("1.0", "end-1c")

    def refresh_batches(self):
        values = directories(INPUT_ROOT)
        self.batch_combo.configure(values=values)
        self.refresh_owners()

    def refresh_owners(self):
        try:
            batch = self.batch_var.get()
            batch = batch if batch in directories(INPUT_ROOT) else clean_task_name(batch)
            values = directories(INPUT_ROOT / batch)
        except ValueError:
            values = []
        self.owner_combo.configure(values=values)
        self.update_task_path()

    def update_task_path(self, *args):
        self.app.vars["link_output"].set("")
        self.path_ready.set("")
        try:
            target = task_path(self.batch_var.get(), self.owner_var.get(), self.task_folder_var.get())
            self.path_preview.set(str(target.relative_to(INPUT_ROOT.parent)))
        except ValueError:
            self.path_preview.set("01INPUT_XLSX/" + "/".join(v for v in (self.batch_var.get(), self.owner_var.get()) if v))

    def prepare_task_folder(self):
        target = task_path(self.batch_var.get(), self.owner_var.get(), self.task_folder_var.get())
        self.batch_var.set(target.parent.parent.name)
        self.owner_var.set(target.parent.name)
        self.task_folder_var.set(target.name)
        existed = target.exists()
        target.mkdir(parents=True, exist_ok=True)
        self.refresh_batches()
        self.app.vars["link_output"].set(str(target))
        self.path_ready.set("使用已有文件夹" if existed else "文件夹已创建")
        return str(target)

    def create_task_folder(self):
        if self.app.busy:
            return
        try:
            self.prepare_task_folder()
        except (ValueError, OSError) as exc:
            messagebox.showerror("创建失败", str(exc), parent=self.app.root)

    def busy(self, state: bool):
        """继承父类禁用逻辑，并同步切换主按钮文案。"""
        super().busy(state)
        self.start_btn.configure(text="运行中…" if state else "开始采集  →")


class ProductPage(BasePage):
    """② 商品采集页：链接表目录选择、文件/分类树与图片规则编辑。"""

    def __init__(self, parent, app: "Workbench"):
        """构建目录选择行、文件与分类树（Treeview）及操作按钮。"""
        super().__init__(parent, app)

        shell = ttk.Frame(self, style="Page.TFrame", padding=(24, 22, 24, 24))
        shell.pack(fill="both", expand=True)

        # 文件目录
        card1, body1 = self.card(shell)
        card1.pack(fill="x")
        self.section_title(body1, "链接表目录", "选择上一步生成的 Excel 链接表目录并读取文件。")

        row = ttk.Frame(body1, style="Card.TFrame")
        row.pack(fill="x", pady=(14, 0))

        self.input_entry = ttk.Entry(
            row,
            textvariable=app.vars["product_input"],
            style="Modern.TEntry",
        )
        self.input_entry.pack(side="left", fill="x", expand=True)

        self.choose_btn = ttk.Button(
            row,
            text="选择目录",
            style="Secondary.TButton",
            command=self.choose_input,
        )
        self.choose_btn.pack(side="left", padx=(8, 0))

        self.scan_btn = ttk.Button(
            row,
            text="读取文件",
            style="Primary.TButton",
            command=app.scan_files,
        )
        self.scan_btn.pack(side="left", padx=(8, 0))

        self.summary_var = tk.StringVar(value="尚未读取文件")
        ttk.Label(body1, textvariable=self.summary_var, style="Muted.TLabel").pack(
            anchor="w", pady=(10, 0)
        )
        self.track_busy(self.input_entry, self.choose_btn, self.scan_btn)

        # 文件列表
        card2, body2 = self.card(shell, padding=0)
        card2.pack(fill="both", expand=True, pady=(16, 0))

        header = ttk.Frame(body2, style="Card.TFrame", padding=(20, 18, 20, 12))
        header.pack(fill="x")
        self.section_title(header, "文件与分类", "双击行可编辑图片保留 / 跳过规则。")

        tree_wrap = ttk.Frame(body2, style="Card.TFrame", padding=(20, 0, 20, 20))
        tree_wrap.pack(fill="both", expand=True)

        columns = ("name", "keep", "skip", "status")
        self.tree = ttk.Treeview(
            tree_wrap,
            columns=columns,
            show="tree headings",
            style="Modern.Treeview",
            selectmode="browse",
        )
        self.tree.heading("#0", text="类型")
        self.tree.heading("name", text="文件 / 分类")
        self.tree.heading("keep", text="保留图片")
        self.tree.heading("skip", text="跳过图片")
        self.tree.heading("status", text="状态")

        self.tree.column("#0", width=78, minwidth=65, anchor="center", stretch=False)
        self.tree.column("name", width=300, minwidth=180, anchor="w")
        self.tree.column("keep", width=150, minwidth=100, anchor="w")
        self.tree.column("skip", width=150, minwidth=100, anchor="w")
        self.tree.column("status", width=100, minwidth=80, anchor="center", stretch=False)

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_wrap.rowconfigure(0, weight=1)
        tree_wrap.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", lambda _e: app.edit_filter())
        self.tree.bind("<Return>", lambda _e: app.edit_filter())

        # 操作区
        actions = ttk.Frame(shell, style="Page.TFrame")
        actions.pack(fill="x", pady=(16, 0))

        self.edit_btn = ttk.Button(
            actions,
            text="编辑图片规则",
            style="Secondary.TButton",
            command=app.edit_filter,
        )
        self.edit_btn.pack(side="left")

        self.next_btn = ttk.Button(
            actions,
            text="继续到数据处理  →",
            style="Secondary.TButton",
            command=lambda: app.handoff("products"),
        )
        self.next_btn.pack(side="right", padx=(8, 0))

        self.start_btn = ttk.Button(
            actions,
            text="开始商品采集  →",
            style="Primary.TButton",
            command=app.start_products,
        )
        self.start_btn.pack(side="right")

        self.track_busy(self.tree, self.edit_btn, self.next_btn, self.start_btn)

    def choose_input(self):
        """弹出目录选择框，选中后写入 product_input（取消则保持原值）。"""
        path = filedialog.askdirectory(
            title="选择链接表目录",
            initialdir=self.app.vars["product_input"].get() or str(APP_DIR.parent),
            parent=self.app.root,
        )
        if path:
            self.app.vars["product_input"].set(path)

    def populate(self, files):
        """按读取结果重建文件树：每个文件一个父节点，其下挂分类子节点。

        节点 iid 约定为 ``f{i}``（文件）与 ``f{i}c{j}``（第 i 个文件的第 j 个
        分类），edit_filter 与 file_status 事件都依赖这个编码定位行。
        """
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        category_total = 0

        for i, info in enumerate(files):
            path = Path(info["path"])
            parent_iid = f"f{i}"

            self.tree.insert(
                "",
                "end",
                iid=parent_iid,
                text="文件",
                values=(
                    path.name,
                    info.get("keep", ""),
                    info.get("skip", ""),
                    "待处理",
                ),
                open=True,
            )

            for j, (title, rule) in enumerate(info.get("categories", {}).items()):
                category_total += 1
                child_iid = f"f{i}c{j}"
                self.tree.insert(
                    parent_iid,
                    "end",
                    iid=child_iid,
                    text="分类",
                    values=(
                        title,
                        rule.get("keep", ""),
                        rule.get("skip", ""),
                        "",
                    ),
                )

        self.summary_var.set(f"已读取 {len(files)} 个文件    {category_total} 个分类")

    def busy(self, state: bool):
        """继承父类禁用逻辑，并同步切换主按钮文案。"""
        super().busy(state)
        self.start_btn.configure(text="运行中…" if state else "开始商品采集  →")


class ProcessingPage(BasePage):
    """③ 数据处理页：中间表目录选择、识别到的品牌展示与结果入口。"""

    def __init__(self, parent, app: "Workbench"):
        """构建「当前识别」品牌横幅、输入目录行、结果卡片与操作按钮。"""
        super().__init__(parent, app)

        shell = ttk.Frame(self, style="Page.TFrame", padding=(24, 22, 24, 24))
        shell.pack(fill="both", expand=True)

        # 当前识别
        info_card = tk.Frame(
            shell,
            bg=COLORS["primary_light"],
            highlightthickness=1,
            highlightbackground="#DBEAFE",
            bd=0,
        )
        info_card.pack(fill="x")

        info_body = tk.Frame(info_card, bg=COLORS["primary_light"], padx=20, pady=18)
        info_body.pack(fill="x")

        tk.Label(
            info_body,
            text="当前识别",
            bg=COLORS["primary_light"],
            fg=COLORS["primary"],
            font=(app.font_family, 10, "bold"),
        ).pack(anchor="w")

        self.brand_display = tk.Label(
            info_body,
            textvariable=app.brand_var,
            bg=COLORS["primary_light"],
            fg=COLORS["text"],
            font=(app.font_family, 13, "bold"),
            anchor="w",
            justify="left",
        )
        self.brand_display.pack(anchor="w", pady=(8, 0))

        # 输入
        card1, body1 = self.card(shell)
        card1.pack(fill="x", pady=(16, 0))
        self.section_title(body1, "输入目录", "选择商品采集阶段生成的中间表目录。")

        row = ttk.Frame(body1, style="Card.TFrame")
        row.pack(fill="x", pady=(14, 0))

        self.input_entry = ttk.Entry(
            row,
            textvariable=app.vars["processing_input"],
            style="Modern.TEntry",
        )
        self.input_entry.pack(side="left", fill="x", expand=True)

        self.choose_btn = ttk.Button(
            row,
            text="选择目录",
            style="Secondary.TButton",
            command=self.choose_input,
        )
        self.choose_btn.pack(side="left", padx=(8, 0))

        self.track_busy(self.input_entry, self.choose_btn)

        # 结果
        self.result_card = tk.Frame(
            shell,
            bg=COLORS["success_bg"],
            highlightthickness=1,
            highlightbackground="#BBF7D0",
            bd=0,
        )
        self.result_card.pack(fill="x", pady=(16, 0))

        result_body = tk.Frame(
            self.result_card,
            bg=COLORS["success_bg"],
            padx=20,
            pady=16,
        )
        result_body.pack(fill="x")

        tk.Label(
            result_body,
            text="✓ 合并结果",
            bg=COLORS["success_bg"],
            fg=COLORS["success"],
            font=(app.font_family, 11, "bold"),
        ).pack(anchor="w")

        tk.Label(
            result_body,
            textvariable=app.result_var,
            bg=COLORS["success_bg"],
            fg=COLORS["text_secondary"],
            font=("TkFixedFont", 9),
            anchor="w",
            justify="left",
            wraplength=760,
        ).pack(anchor="w", pady=(8, 0))

        result_actions = tk.Frame(result_body, bg=COLORS["success_bg"])
        result_actions.pack(fill="x", pady=(12, 0))

        self.open_btn = ttk.Button(
            result_actions,
            text="打开结果目录",
            style="Secondary.TButton",
            command=app.open_results,
        )
        self.open_btn.pack(side="right")

        # 操作
        actions = ttk.Frame(shell, style="Page.TFrame")
        actions.pack(fill="x", pady=(16, 0))

        self.start_btn = ttk.Button(
            actions,
            text="开始规范与合并  →",
            style="Primary.TButton",
            command=app.start_processing,
        )
        self.start_btn.pack(side="right")

        self.handoff_btn = ttk.Button(actions, text="继续到转换与导出 →", style="Secondary.TButton", command=lambda: app.handoff("processing"))
        self.handoff_btn.pack(side="right", padx=(0, 12))
        self.track_busy(self.start_btn, self.open_btn, self.handoff_btn)

    def choose_input(self):
        """弹出目录选择框，选中后写入 processing_input。

        该变量带 trace_add 监听，赋值后会自动刷新顶部「当前识别」的品牌。
        """
        path = filedialog.askdirectory(
            title="选择中间表目录",
            initialdir=self.app.vars["processing_input"].get() or str(APP_DIR.parent),
            parent=self.app.root,
        )
        if path:
            self.app.vars["processing_input"].set(path)

    def busy(self, state: bool):
        """继承父类禁用逻辑，并同步切换主按钮文案。"""
        super().busy(state)
        self.start_btn.configure(text="运行中…" if state else "开始规范与合并  →")


class ConversionPage(BasePage):
    """④ 转换与导出页，独立选择已经规范好的合并表。"""
    def __init__(self, parent, app):
        super().__init__(parent, app)
        shell = ttk.Frame(self, style="Page.TFrame", padding=24)
        shell.pack(fill="both", expand=True)
        card, body = self.card(shell)
        card.pack(fill="x")
        self.section_title(body, "合并表")
        row = ttk.Frame(body, style="Card.TFrame")
        row.pack(fill="x", pady=14)
        entry = ttk.Entry(row, textvariable=app.vars["conversion_input"], style="Modern.TEntry")
        entry.pack(side="left", fill="x", expand=True)
        choose = ttk.Button(row, text="选择文件", style="Secondary.TButton", command=self.choose_input)
        choose.pack(side="left", padx=(8, 0))
        ttk.Label(body, textvariable=app.conversion_result_var, style="Muted.TLabel", wraplength=760).pack(anchor="w")
        actions = ttk.Frame(shell, style="Page.TFrame")
        actions.pack(fill="x", pady=16)
        start = ttk.Button(actions, text="开始转换与导出 →", style="Primary.TButton", command=app.start_conversion)
        start.pack(side="right")
        open_button = ttk.Button(actions, text="打开结果目录", style="Secondary.TButton", command=lambda: app.open_results("conversion"))
        open_button.pack(side="right", padx=12)
        self.track_busy(entry, choose, start, open_button)

    def choose_input(self):
        path = filedialog.askopenfilename(parent=self.app.root, title="选择合并表", filetypes=[("Excel 合并表", "*.xlsx")])
        if path:
            self.app.vars["conversion_input"].set(path)


class Workbench:
    """主窗口控制器：构建界面、持有各阶段输入变量、调度任务并在主线程消费事件。"""

    def __init__(self, root):
        """初始化默认设置、阶段状态变量，构建界面并启动事件轮询。"""
        self.root = root
        self.controller = TaskController()
        self.busy = False
        self.closing = False
        self.files = []
        self.outputs = {}
        self.logs = []
        self.page_key = "links"

        defaults = {
            "proxy": "http://127.0.0.1:7897",
            "link_output": str(APP_DIR.parent / "01INPUT_XLSX"),
            "product_input": "",
            "processing_input": "",
            "conversion_input": "",
            "max_pages": "",
            "xpath": "",
        }
        self.vars = {key: tk.StringVar(value=value) for key, value in defaults.items()}

        self.status_var = tk.StringVar(value="就绪")
        self.count_var = tk.StringVar(value="")
        self.brand_var = tk.StringVar(value="目标品牌：—")
        self.result_var = tk.StringVar(value="")
        self.conversion_result_var = tk.StringVar(value="")
        self.merged_file = None
        self.stage_states = {
            key: tk.StringVar(value="待运行")
            for key in ("links", "products", "processing", "conversion")
        }
        self.log_filter = tk.StringVar(value="全部")
        self.error_only = tk.BooleanVar()

        self.header_controls = []
        self.nav_buttons = {}
        self.nav_badges = {}
        self.page_title_var = tk.StringVar()
        self.page_subtitle_var = tk.StringVar()

        self.font_family = choose_font(root)

        self._build()
        root.protocol("WM_DELETE_WINDOW", self.close)
        self.vars["processing_input"].trace_add("write", self.update_brand)
        self._poll_id = root.after(100, self.poll)

    # ------------------------------------------------------------------
    # 构建 UI
    # ------------------------------------------------------------------

    def _build(self):
        """设置窗口基本属性，依次装配 ttk 样式、页面骨架，并默认停在阶段①。"""
        root = self.root
        root.title("商品采集与数据处理工作台")
        root.geometry("1280x820")
        root.minsize(1080, 680)
        root.configure(bg=COLORS["bg"])

        self._setup_styles()
        self._build_shell()
        self.show("links")

    def _setup_styles(self):
        """集中注册全部 ttk 样式（字体、配色、按钮/输入框/表格/进度条）。"""
        style = ttk.Style(self.root)
        if "clam" in style.theme_names():
            style.theme_use("clam")

        family = self.font_family

        style.configure(".", font=(family, 10))
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Page.TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["surface"])

        style.configure(
            "TLabel",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=(family, 10),
        )
        style.configure(
            "TopbarTitle.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(family, 20, "bold"),
        )
        style.configure(
            "TopbarSubtitle.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text_secondary"],
            font=(family, 10),
        )
        style.configure(
            "CardTitle.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(family, 12, "bold"),
        )
        style.configure(
            "FieldLabel.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=(family, 10, "bold"),
        )
        style.configure(
            "Hint.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text_muted"],
            font=(family, 9),
        )
        style.configure(
            "Muted.TLabel",
            background=COLORS["surface"],
            foreground=COLORS["text_secondary"],
            font=(family, 9),
        )

        # Entry
        style.configure(
            "Modern.TEntry",
            padding=(10, 8),
            relief="flat",
            fieldbackground=COLORS["surface_soft"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border_strong"],
            lightcolor=COLORS["border_strong"],
            darkcolor=COLORS["border_strong"],
        )
        style.map(
            "Modern.TEntry",
            bordercolor=[("focus", COLORS["primary"])],
            lightcolor=[("focus", COLORS["primary"])],
            darkcolor=[("focus", COLORS["primary"])],
        )

        # Buttons
        style.configure(
            "Primary.TButton",
            padding=(14, 9),
            background=COLORS["primary"],
            foreground="#FFFFFF",
            borderwidth=0,
            relief="flat",
            font=(family, 10, "bold"),
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", "#CBD5E1"),
                ("active", COLORS["primary_hover"]),
                ("pressed", COLORS["primary_hover"]),
            ],
            foreground=[("disabled", "#F8FAFC"), ("!disabled", "#FFFFFF")],
        )

        style.configure(
            "Secondary.TButton",
            padding=(14, 9),
            background=COLORS["surface"],
            foreground=COLORS["text"],
            borderwidth=1,
            relief="flat",
            font=(family, 10),
        )
        style.map(
            "Secondary.TButton",
            background=[
                ("active", COLORS["surface_soft"]),
                ("pressed", "#F1F5F9"),
                ("disabled", "#F8FAFC"),
            ],
            foreground=[("disabled", COLORS["text_muted"])],
        )

        style.configure(
            "Ghost.TButton",
            padding=(10, 8),
            background=COLORS["surface"],
            foreground=COLORS["text_secondary"],
            borderwidth=0,
            relief="flat",
        )
        style.map(
            "Ghost.TButton",
            background=[("active", COLORS["surface_soft"])],
            foreground=[("active", COLORS["text"])],
        )

        style.configure(
            "Danger.TButton",
            padding=(12, 8),
            background=COLORS["danger_bg"],
            foreground=COLORS["danger"],
            borderwidth=0,
            relief="flat",
            font=(family, 10, "bold"),
        )
        style.map(
            "Danger.TButton",
            background=[
                ("active", "#FEE2E2"),
                ("disabled", "#F8FAFC"),
            ],
            foreground=[("disabled", COLORS["text_muted"])],
        )

        # Treeview
        style.configure(
            "Modern.Treeview",
            background=COLORS["surface"],
            fieldbackground=COLORS["surface"],
            foreground=COLORS["text"],
            rowheight=34,
            relief="flat",
            borderwidth=0,
            font=(family, 9),
        )
        style.configure(
            "Modern.Treeview.Heading",
            background=COLORS["surface_soft"],
            foreground=COLORS["text_secondary"],
            relief="flat",
            font=(family, 9, "bold"),
            padding=(8, 8),
        )
        style.map(
            "Modern.Treeview",
            background=[("selected", COLORS["primary_light"])],
            foreground=[("selected", COLORS["text"])],
        )
        style.map(
            "Modern.Treeview.Heading",
            background=[("active", "#F1F5F9")],
        )

        # Combobox
        style.configure(
            "Modern.TCombobox",
            padding=(8, 6),
            fieldbackground=COLORS["surface"],
            foreground=COLORS["text"],
        )

        # Progressbar
        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor="#E2E8F0",
            background=COLORS["primary"],
            borderwidth=0,
            thickness=5,
        )

    def _build_shell(self):
        """装配整体骨架：左侧导航栏 + 右侧（顶栏 / 页面容器 / 底部状态栏）。

        所有页面共用 content 的同一个 grid 单元，切换页面靠 tkraise 实现。
        """
        root = self.root

        shell = tk.Frame(root, bg=COLORS["bg"])
        shell.pack(fill="both", expand=True)

        sidebar = self._build_sidebar(shell)

        main = tk.Frame(shell, bg=COLORS["bg"])
        main.pack(side="left", fill="both", expand=True)

        self._build_topbar(main)

        content = ttk.Frame(main, style="Page.TFrame")
        content.pack(fill="both", expand=True)
        content.rowconfigure(0, weight=1)
        content.columnconfigure(0, weight=1)

        self.pages = {
            "links": LinkPage(content, self),
            "products": ProductPage(content, self),
            "processing": ProcessingPage(content, self),
            "conversion": ConversionPage(content, self),
        }
        self.pages["logs"] = self._build_log_page(content)

        for page in self.pages.values():
            page.grid(row=0, column=0, sticky="nsew")

        self._build_footer(main)

    def _build_sidebar(self, parent):
        """构建深色侧边栏：品牌区、四个阶段导航项、日志入口与底部版本信息。"""
        sidebar = tk.Frame(parent, bg=COLORS["sidebar"], width=210)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        brand = tk.Frame(sidebar, bg=COLORS["sidebar"], padx=20, pady=22)
        brand.pack(fill="x")

        tk.Label(
            brand,
            text="◈  DataFlow",
            bg=COLORS["sidebar"],
            fg="#FFFFFF",
            font=(self.font_family, 15, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            brand,
            text="商品采集 · 清洗 · 处理",
            bg=COLORS["sidebar"],
            fg=COLORS["text_muted"],
            font=(self.font_family, 9),
            anchor="w",
        ).pack(fill="x", pady=(5, 0))

        divider = tk.Frame(sidebar, bg="#253044", height=1)
        divider.pack(fill="x", padx=16)

        nav = tk.Frame(sidebar, bg=COLORS["sidebar"], padx=10, pady=16)
        nav.pack(fill="x")

        entries = [
            ("links", "01", "链接采集"),
            ("products", "02", "商品采集"),
            ("processing", "03", "规范与合并"),
            ("conversion", "04", "转换与导出"),
        ]

        for key, number, label in entries:
            self._create_nav_item(nav, key, number, label, with_badge=True)

        sep = tk.Frame(nav, bg="#253044", height=1)
        sep.pack(fill="x", padx=8, pady=(10, 12))

        self._create_nav_item(nav, "logs", "", "运行日志", with_badge=False)

        version_wrap = tk.Frame(sidebar, bg=COLORS["sidebar"], padx=20, pady=16)
        version_wrap.pack(side="bottom", fill="x")
        tk.Label(
            version_wrap,
            text="Desktop Workbench",
            bg=COLORS["sidebar"],
            fg="#64748B",
            font=(self.font_family, 8),
            anchor="w",
        ).pack(fill="x")

        return sidebar

    def _create_nav_item(self, parent, key, number, label, with_badge):
        """创建一个导航项，并登记其子控件供选中/悬停状态统一改色。

        阶段项额外挂一个状态徽标，绑定 stage_states 变化自动刷新文案与颜色。
        """
        item = tk.Frame(parent, bg=COLORS["sidebar"], height=46, cursor="hand2")
        item.pack(fill="x", pady=3)
        item.pack_propagate(False)

        left = tk.Frame(item, bg=COLORS["sidebar"])
        left.pack(side="left", fill="y", padx=(12, 0))

        if number:
            number_label = tk.Label(
                left,
                text=number,
                bg=COLORS["sidebar"],
                fg="#64748B",
                font=("TkFixedFont", 9, "bold"),
            )
            number_label.pack(side="left", pady=13, padx=(0, 10))
        else:
            number_label = None

        label_widget = tk.Label(
            left,
            text=label,
            bg=COLORS["sidebar"],
            fg="#CBD5E1",
            font=(self.font_family, 10, "bold"),
        )
        label_widget.pack(side="left", pady=12)

        badge = None
        if with_badge:
            badge = tk.Label(
                item,
                text=f"● {self.stage_states[key].get()}",
                bg=COLORS["sidebar"],
                fg=COLORS["text_muted"],
                font=(self.font_family, 8),
            )
            badge.pack(side="right", padx=(0, 10))
            self.nav_badges[key] = badge
            self.stage_states[key].trace_add(
                "write",
                lambda *_args, k=key: self._update_stage_badge(k),
            )

        widgets = [item, left, label_widget]
        if number_label:
            widgets.append(number_label)
        if badge:
            widgets.append(badge)

        for widget in widgets:
            widget.bind("<Button-1>", lambda _e, k=key: self.show(k))
            widget.bind("<Enter>", lambda _e, k=key: self._nav_hover(k, True))
            widget.bind("<Leave>", lambda _e, k=key: self._nav_hover(k, False))

        self.nav_buttons[key] = {
            "frame": item,
            "left": left,
            "label": label_widget,
            "number": number_label,
            "badge": badge,
        }

    def _build_topbar(self, parent):
        """构建顶栏：左侧当前页标题/副标题，右侧浏览器代理输入框，底部一条分隔线。"""
        topbar = tk.Frame(
            parent,
            bg=COLORS["surface"],
            height=76,
            highlightthickness=0,
        )
        topbar.pack(fill="x")
        topbar.pack_propagate(False)

        left = tk.Frame(topbar, bg=COLORS["surface"])
        left.pack(side="left", fill="both", expand=True, padx=(24, 12), pady=(14, 10))

        self.page_title_label = tk.Label(
            left,
            textvariable=self.page_title_var,
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=(self.font_family, 18, "bold"),
            anchor="w",
        )
        self.page_title_label.pack(anchor="w")

        self.page_subtitle_label = tk.Label(
            left,
            textvariable=self.page_subtitle_var,
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9),
            anchor="w",
        )
        self.page_subtitle_label.pack(anchor="w", pady=(3, 0))

        right = tk.Frame(topbar, bg=COLORS["surface"])
        right.pack(side="right", padx=(12, 24), pady=15)

        tk.Label(
            right,
            text="浏览器代理",
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9),
        ).pack(side="left", padx=(0, 8))

        proxy_entry = ttk.Entry(
            right,
            textvariable=self.vars["proxy"],
            width=28,
            style="Modern.TEntry",
        )
        proxy_entry.pack(side="left")
        self.header_controls.append(proxy_entry)

        line = tk.Frame(parent, bg=COLORS["border"], height=1)
        line.pack(fill="x")

    def _build_log_page(self, parent):
        """构建日志页：阶段筛选下拉、仅错误/警告开关与深色终端风格文本框。

        文本框 predefine 了 normal/success/warning/error/stage 五个 tag，
        供 append_log 按内容着色。
        """
        page = ttk.Frame(parent, style="Page.TFrame")
        shell = ttk.Frame(page, style="Page.TFrame", padding=(24, 22, 24, 24))
        shell.pack(fill="both", expand=True)

        card = tk.Frame(
            shell,
            bg=COLORS["surface"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            bd=0,
        )
        card.pack(fill="both", expand=True)

        header = tk.Frame(card, bg=COLORS["surface"], padx=20, pady=16)
        header.pack(fill="x")

        tools = tk.Frame(header, bg=COLORS["surface"])
        tools.pack(side="right")

        tk.Label(
            tools,
            text="阶段",
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9),
        ).pack(side="left", padx=(0, 6))

        combo = ttk.Combobox(
            tools,
            textvariable=self.log_filter,
            values=["全部", "links", "products", "processing", "conversion", "scan"],
            state="readonly",
            width=14,
            style="Modern.TCombobox",
        )
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", lambda _e: self.render_logs())

        error_check = ttk.Checkbutton(
            tools,
            text="仅错误 / 警告",
            variable=self.error_only,
            command=self.render_logs,
        )
        error_check.pack(side="left", padx=(12, 0))

        console_wrap = tk.Frame(card, bg=COLORS["console"], padx=0, pady=0)
        console_wrap.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.log_text = scrolledtext.ScrolledText(
            console_wrap,
            state="disabled",
            bg=COLORS["console"],
            fg=COLORS["console_text"],
            insertbackground="#FFFFFF",
            wrap="word",
            relief="flat",
            bd=0,
            padx=14,
            pady=14,
            font=("TkFixedFont", 10),
        )
        self.log_text.pack(fill="both", expand=True)

        self.log_text.tag_configure("normal", foreground=COLORS["console_text"])
        self.log_text.tag_configure("success", foreground="#4ADE80")
        self.log_text.tag_configure("warning", foreground="#FBBF24")
        self.log_text.tag_configure("error", foreground="#F87171")
        self.log_text.tag_configure("stage", foreground="#93C5FD")

        return page

    def _build_footer(self, parent):
        """构建底部状态栏：进度条、状态圆点与文案、计数、停止按钮。"""
        footer = tk.Frame(
            parent,
            bg=COLORS["surface"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
            height=72,
        )
        footer.pack(side="bottom", fill="x")
        footer.pack_propagate(False)

        self.progress = ttk.Progressbar(
            footer,
            mode="determinate",
            maximum=100,
            style="Modern.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", side="top")

        body = tk.Frame(footer, bg=COLORS["surface"])
        body.pack(fill="both", expand=True, padx=24, pady=12)

        left = tk.Frame(body, bg=COLORS["surface"])
        left.pack(side="left", fill="both", expand=True)

        self.status_dot = tk.Label(
            left,
            text="●",
            bg=COLORS["surface"],
            fg=COLORS["success"],
            font=(self.font_family, 9),
        )
        self.status_dot.pack(side="left")

        self.status_label = tk.Label(
            left,
            textvariable=self.status_var,
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9),
            anchor="w",
        )
        self.status_label.pack(side="left", padx=(7, 0))

        right = tk.Frame(body, bg=COLORS["surface"])
        right.pack(side="right")

        tk.Label(
            right,
            textvariable=self.count_var,
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9, "bold"),
        ).pack(side="left", padx=(0, 16))

        self.stop_btn = ttk.Button(
            right,
            text="停止任务",
            style="Danger.TButton",
            command=self.stop,
            state="disabled",
        )
        self.stop_btn.pack(side="left")

    # ------------------------------------------------------------------
    # UI 状态
    # ------------------------------------------------------------------

    def show(self, key):
        """切换到指定页面，并同步顶栏标题副标题与侧栏选中态。"""
        self.page_key = key
        self.pages[key].tkraise()

        title, subtitle = PAGE_META[key]
        self.page_title_var.set(title)
        self.page_subtitle_var.set(subtitle)

        self._update_nav_state()

    def _update_nav_state(self):
        """按当前页高亮对应导航项，其余恢复常态配色。"""
        for key, item in self.nav_buttons.items():
            active = key == self.page_key
            bg = COLORS["sidebar_active"] if active else COLORS["sidebar"]

            item["frame"].configure(bg=bg)
            item["left"].configure(bg=bg)
            item["label"].configure(
                bg=bg,
                fg="#FFFFFF" if active else "#CBD5E1",
            )
            if item["number"]:
                item["number"].configure(
                    bg=bg,
                    fg="#DBEAFE" if active else "#64748B",
                )
            if item["badge"]:
                item["badge"].configure(bg=bg)

    def _nav_hover(self, key, entering):
        """鼠标进出导航项时切换底色；当前页不参与悬停变色。"""
        if key == self.page_key:
            return
        item = self.nav_buttons[key]
        bg = COLORS["sidebar_hover"] if entering else COLORS["sidebar"]
        item["frame"].configure(bg=bg)
        item["left"].configure(bg=bg)
        item["label"].configure(bg=bg)
        if item["number"]:
            item["number"].configure(bg=bg)
        if item["badge"]:
            item["badge"].configure(bg=bg)

    def _update_stage_badge(self, key):
        """按阶段状态文案刷新侧栏徽标的文本与颜色（未知状态回落灰色）。"""
        badge = self.nav_badges.get(key)
        if not badge:
            return

        state = self.stage_states[key].get()
        color_map = {
            "待运行": COLORS["text_muted"],
            "运行中": "#60A5FA",
            "完成": "#4ADE80",
            "失败": "#F87171",
            "部分失败": "#FBBF24",
            "已停止": "#FBBF24",
        }
        badge.configure(
            text=f"● {state}",
            fg=color_map.get(state, COLORS["text_muted"]),
        )

    def _set_busy_visual(self, busy: bool):
        """状态圆点：运行中蓝色，空闲绿色。"""
        self.status_dot.configure(
            fg=COLORS["primary"] if busy else COLORS["success"]
        )

    # ------------------------------------------------------------------
    # 原业务逻辑
    # ------------------------------------------------------------------

    def update_brand(self, *_args):
        """输入目录变化时，从目录名解析并展示目标域名与品牌；无法识别则提示未识别。"""
        try:
            brand = extract_brand(Path(self.vars["processing_input"].get()).name)
            self.brand_var.set(
                f"目标域名：{brand.domain}    目标品牌：{brand.word}"
            )
        except ValueError:
            self.brand_var.set("目标品牌：未识别")

    def settings(self):
        """把所有输入框值去除首尾空白后打包成任务参数快照。"""
        return {
            key: value.get().strip()
            for key, value in self.vars.items()
        }

    def launch(self, stage, operation):
        """启动一个后台任务并进入「运行中」界面状态。

        启动失败（通常是已有任务在跑）只弹错误框并保持原状；
        成功后禁用全部输入控件、点亮停止按钮并重置进度。
        """
        if self.busy:
            return

        try:
            self.controller.start(stage, operation)
        except Exception as exc:
            messagebox.showerror("启动失败", str(exc), parent=self.root)
            return

        self.busy = True

        for key, page in self.pages.items():
            if key != "logs" and hasattr(page, "busy"):
                page.busy(True)

        for widget in self.header_controls:
            widget.configure(state="disabled")

        self.stop_btn.configure(state="normal")
        self.status_var.set("正在读取文件…" if stage == "scan" else "正在运行…")
        self.count_var.set("")
        self.progress.configure(value=0, maximum=100)
        self._set_busy_visual(True)

        if stage in self.stage_states:
            self.stage_states[stage].set("运行中")
            self.outputs.pop(stage, None)

            if stage == "processing":
                self.result_var.set("")
                self.merged_file = None
            if stage == "conversion":
                self.conversion_result_var.set("")

    def start_links(self):
        """校验分类文本/输出目录/页数后启动第一阶段任务。"""
        if self.busy:
            return

        settings = self.settings()
        # 分类文本取文本框全部内容（去掉 Tk 自动附加的末尾换行）
        settings["categories"] = self.pages["links"].category_input()

        try:
            if not parse_categories(settings["categories"]):
                raise ValueError("请输入有效分类 URL")

            raw = settings["max_pages"]
            settings["max_pages"] = int(raw) if raw else 0

            if raw and settings["max_pages"] <= 0:
                raise ValueError("页数应为正整数，留空不限")

        except ValueError as exc:
            messagebox.showerror("输入无效", str(exc), parent=self.root)
            return

        try:
            settings["link_output"] = self.pages["links"].prepare_task_folder()
        except (ValueError, OSError) as exc:
            messagebox.showerror("目录无效", str(exc), parent=self.root)
            return
        self.launch("links", collect_links(settings))

    def scan_files(self):
        """扫描链接表目录，读取每个 xlsx 的分类列表并刷新文件树。

        已存在的图片规则按路径复用（previous 快照），重扫不会丢失手工配置；
        恢复文件（含 .recovery-）只提示并跳过。整个过程作为 scan 任务在后台跑，
        因为逐个读表在大目录下会明显卡界面。
        """
        if self.busy:
            return

        raw = self.vars["product_input"].get().strip()
        if not raw or not Path(raw).is_dir():
            messagebox.showerror(
                "路径无效",
                "请选择有效的链接表目录",
                parent=self.root,
            )
            return

        folder = Path(raw).resolve()
        previous = {f["path"]: f for f in copy.deepcopy(self.files)}

        def operation(controller, log):
            """后台读取目录：产出 files 列表与目录路径，供 poll 回填界面。"""
            files = []

            for path in sorted(folder.iterdir()):
                checkpoint(controller.stop_event)

                if (
                    not path.is_file()
                    or path.suffix.lower() != ".xlsx"
                    or path.name.startswith(("~", "."))
                ):
                    continue

                if ".recovery-" in path.name:
                    log(f"跳过恢复文件：{path.name}")
                    continue

                cats = read_categories_from_excel(path)
                old = previous.get(str(path), {})
                rules = old.get("categories", {})

                files.append(
                    {
                        "path": str(path),
                        "keep": old.get("keep", ""),
                        "skip": old.get("skip", ""),
                        "categories": {
                            title: rules.get(
                                title,
                                {"keep": "", "skip": ""},
                            )
                            for title in cats
                        },
                    }
                )

                log(f"{path.name}：{len(cats)} 个分类")

            if not files:
                raise ValueError("没有可读取的链接表")

            return {
                "files": files,
                "folder": str(folder),
                "status": "完成",
            }

        self.launch("scan", operation)

    def edit_filter(self):
        """弹出图片规则对话框，编辑选中行的保留/跳过设置。

        选中行由 iid 编码定位：``f{i}`` 编辑文件级默认规则，
        ``f{i}c{j}`` 编辑该文件第 j 个分类的独立规则；
        保存后同时写回 self.files 与文件树单元格。
        """
        if self.busy:
            return

        tree = self.pages["products"].tree
        selection = tree.selection()

        if not selection:
            messagebox.showinfo(
                "未选择",
                "请先选择一个文件或分类。",
                parent=self.root,
            )
            return

        iid = selection[0]
        parts = iid[1:].split("c")
        info = self.files[int(parts[0])]

        # 长度 1 表示文件行，否则按分类序号取对应规则字典
        rule = (
            info
            if len(parts) == 1
            else list(info["categories"].values())[int(parts[1])]
        )

        dialog = tk.Toplevel(self.root)
        dialog.title("图片过滤规则")
        dialog.geometry("460x270")
        dialog.minsize(430, 250)
        dialog.configure(bg=COLORS["bg"])
        dialog.transient(self.root)
        dialog.grab_set()

        # 居中
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 460) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 270) // 2
        dialog.geometry(f"+{max(x, 0)}+{max(y, 0)}")

        card = tk.Frame(
            dialog,
            bg=COLORS["surface"],
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        card.pack(fill="both", expand=True, padx=18, pady=18)

        body = tk.Frame(card, bg=COLORS["surface"], padx=20, pady=18)
        body.pack(fill="both", expand=True)

        tk.Label(
            body,
            text="图片过滤规则",
            bg=COLORS["surface"],
            fg=COLORS["text"],
            font=(self.font_family, 14, "bold"),
        ).pack(anchor="w")

        tk.Label(
            body,
            text="设置需要保留或跳过的图片关键词。",
            bg=COLORS["surface"],
            fg=COLORS["text_secondary"],
            font=(self.font_family, 9),
        ).pack(anchor="w", pady=(5, 14))

        values = {
            key: tk.StringVar(value=rule[key])
            for key in ("keep", "skip")
        }

        form = tk.Frame(body, bg=COLORS["surface"])
        form.pack(fill="x")

        for i, (key, label) in enumerate(
            [("keep", "保留图片"), ("skip", "跳过图片")]
        ):
            tk.Label(
                form,
                text=label,
                bg=COLORS["surface"],
                fg=COLORS["text"],
                font=(self.font_family, 9, "bold"),
            ).grid(row=i, column=0, sticky="w", pady=8, padx=(0, 10))

            entry = ttk.Entry(
                form,
                textvariable=values[key],
                style="Modern.TEntry",
            )
            entry.grid(row=i, column=1, sticky="ew", pady=8)

        form.columnconfigure(1, weight=1)

        buttons = tk.Frame(body, bg=COLORS["surface"])
        buttons.pack(fill="x", pady=(16, 0))

        ttk.Button(
            buttons,
            text="取消",
            style="Secondary.TButton",
            command=dialog.destroy,
        ).pack(side="right", padx=(8, 0))

        def apply():
            """把对话框输入写回规则字典，并同步刷新树上该行的两列显示。"""
            rule.update(
                {
                    key: value.get().strip()
                    for key, value in values.items()
                }
            )
            tree.set(iid, "keep", rule["keep"])
            tree.set(iid, "skip", rule["skip"])
            dialog.destroy()

        ttk.Button(
            buttons,
            text="保存规则",
            style="Primary.TButton",
            command=apply,
        ).pack(side="right")

    def start_products(self):
        """确认当前文件列表确实来自所选目录后，启动第二阶段采集。"""
        if self.busy:
            return

        settings = self.settings()

        # 目录被改过而文件树还是旧目录的读取结果时拒绝启动，避免采错数据
        if (
            not self.files
            or any(
                Path(f["path"]).parent
                != Path(settings["product_input"]).resolve()
                for f in self.files
            )
        ):
            messagebox.showerror(
                "需要读取文件",
                "请先读取当前链接表目录",
                parent=self.root,
            )
            return

        self.launch(
            "products",
            collect_products(settings, copy.deepcopy(self.files)),
        )

    def start_processing(self):
        """校验中间表目录（含域名可识别）后，启动第三阶段规范与合并。"""
        if self.busy:
            return

        settings = self.settings()

        try:
            if (
                not settings["processing_input"]
                or not Path(settings["processing_input"]).is_dir()
            ):
                raise ValueError("请选择有效的中间表目录")

            # 提前校验目录末级名能否解析出唯一域名，避免任务跑一半才失败
            extract_brand(Path(settings["processing_input"]).resolve().name)

        except ValueError as exc:
            messagebox.showerror(
                "输入无效",
                str(exc),
                parent=self.root,
            )
            return

        self.launch("processing", process_data(settings))

    def start_conversion(self):
        if self.busy:
            return
        settings = self.settings()
        source = Path(settings["conversion_input"])
        if not source.is_file() or source.suffix.lower() != ".xlsx":
            messagebox.showerror("输入无效", "请选择有效的 XLSX 合并表", parent=self.root)
            return
        self.launch("conversion", convert_data(settings))

    def handoff(self, stage):
        """把某阶段的输出目录交给下一阶段。

        第一阶段会顺带切页并读取文件列表（但不自动开始采集）；
        第二阶段只填路径并切页，等用户确认后再点开始。
        """
        if self.busy:
            return

        if stage == "processing":
            if self.merged_file and Path(self.merged_file).is_file():
                self.vars["conversion_input"].set(self.merged_file)
                self.show("conversion")
            else:
                messagebox.showinfo("暂无合并表", "请先完成规范与合并", parent=self.root)
            return
        output = self.outputs.get(stage)

        if not output or not Path(output).is_dir():
            messagebox.showinfo(
                "暂无输出",
                "请先完成该阶段，或在下一阶段手动选择已有数据目录",
                parent=self.root,
            )
            return

        if stage == "links":
            self.vars["product_input"].set(output)
            self.show("products")
            self.scan_files()
        else:
            self.vars["processing_input"].set(output)
            self.show("processing")

    def stop(self):
        """请求停止当前任务，并禁用停止按钮避免重复点击。"""
        if self.busy:
            self.controller.stop()
            self.stop_btn.configure(state="disabled")
            self.status_var.set("正在停止并等待保存…")

    # ------------------------------------------------------------------
    # 事件轮询
    # ------------------------------------------------------------------

    def poll(self):
        """主线程事件泵：每 100ms 消费一批后台事件并更新界面。

        每次最多处理 250 条，避免日志爆发时长时间阻塞 UI；
        done 事件若早于线程退出到达，会放回队列等下一轮再处理；
        closing 且任务已结束时执行真正的退出。
        """
        for _ in range(250):
            try:
                kind, value = self.controller.events.get_nowait()
            except queue.Empty:
                break

            if kind == "log":
                self.logs.append(value)

                if len(self.logs) > 10000:
                    self.logs = self.logs[-8000:]
                    self.render_logs()
                else:
                    self.append_log(*value)

            elif kind == "progress":
                current, total, text = value
                self.progress.configure(
                    maximum=max(total, 1),
                    value=current,
                )
                if not self.controller.stop_event.is_set():
                    self.status_var.set(text)

            elif kind == "count":
                self.count_var.set(value)

            elif kind == "file_status":
                i, text = value
                tree = self.pages["products"].tree
                if tree.exists(f"f{i}"):
                    tree.set(f"f{i}", "status", text)

            elif kind == "done":
                # 事件可能先于线程 return，等线程结束再启用下一次运行。
                if self.controller.running:
                    self.controller.events.put((kind, value))
                    break

                stage, status, result = value
                self.busy = False

                for key, page in self.pages.items():
                    if key != "logs" and hasattr(page, "busy"):
                        page.busy(False)

                for widget in self.header_controls:
                    widget.configure(state="normal")

                self.stop_btn.configure(state="disabled")
                self._set_busy_visual(False)

                if stage in self.stage_states:
                    self.stage_states[stage].set(status)

                self.status_var.set(status)

                if result:
                    if stage == "scan":
                        self.files = result["files"]
                        self.vars["product_input"].set(result["folder"])
                        self.pages["products"].populate(self.files)

                    elif result.get("output"):
                        self.outputs[stage] = result["output"]
                        self.count_var.set(
                            f"{result.get('count', 0)} 条数据"
                        )

                        if stage == "processing":
                            self.merged_file = result["merged_file"]
                            self.result_var.set(self.merged_file)
                        elif stage == "conversion":
                            self.conversion_result_var.set(result["output"])

                if status == "完成":
                    self.progress.configure(
                        value=self.progress["maximum"]
                    )

                if status in ("失败", "部分失败"):
                    self.closing = False
                    self.show("logs")

        if self.closing and not self.busy:
            self.shutdown()
            return

        self._poll_id = self.root.after(100, self.poll)

    # ------------------------------------------------------------------
    # 日志
    # ------------------------------------------------------------------

    def _log_tag(self, text: str) -> str:
        """按关键词把一行日志归类到 error/warning/success/normal 四个着色 tag。

        判定顺序即优先级：失败类关键词先命中，其次警告类，最后成功类。
        """
        lower = text.lower()

        if any(
            word in text
            for word in ("失败", "异常", "错误", "Traceback", "Error")
        ):
            return "error"

        if any(
            word in text
            for word in ("警告", "恢复", "保护")
        ):
            return "warning"

        if any(
            word in text
            for word in ("完成", "成功", "已保存", "✓")
        ):
            return "success"

        if "warning" in lower:
            return "warning"

        return "normal"

    def append_log(self, stage, text):
        """向日志框追加一行（带阶段前缀并按内容着色）。

        受阶段筛选与「仅错误/警告」两个开关约束，被过滤的行直接丢弃。
        """
        if self.log_filter.get() not in ("全部", stage):
            return

        if self.error_only.get() and not any(
            word in text
            for word in (
                "失败",
                "异常",
                "错误",
                "警告",
                "Error",
                "Traceback",
                "恢复",
                "保护",
            )
        ):
            return

        tag = self._log_tag(text)

        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stage}] ", "stage")
        self.log_text.insert("end", f"{text}\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def render_logs(self):
        """按当前筛选条件清空并重绘整份日志（切换筛选或日志被裁剪时调用）。"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        for stage, text in self.logs:
            self.append_log(stage, text)

    # ------------------------------------------------------------------
    # 结果与退出
    # ------------------------------------------------------------------

    def open_results(self, stage="processing"):
        """打开第三阶段结果目录；尚未产出结果时提示。"""
        output = self.outputs.get(stage)

        if output:
            self.open_directory(Path(output))
        else:
            messagebox.showinfo(
                "暂无结果",
                "尚无处理结果",
                parent=self.root,
            )

    def open_directory(self, path):
        """在系统文件管理器中打开目录（按平台选择 open/startfile/xdg-open）。"""
        if not path.is_dir():
            messagebox.showinfo(
                "目录不存在",
                str(path),
                parent=self.root,
            )
            return

        if sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        elif os.name == "nt":
            os.startfile(path)
        else:
            subprocess.Popen(["xdg-open", str(path)])

    def close(self):
        """窗口关闭入口：有任务在跑则先请求停止，等收尾完成后再真正退出。"""
        if self.busy:
            self.closing = True
            self.stop()
        else:
            self.shutdown()

    def shutdown(self):
        """取消轮询定时器并销毁窗口；任务仍在跑时退回 close 等待收尾。"""
        if self.busy:
            self.close()
            return

        if self._poll_id:
            self.root.after_cancel(self._poll_id)
            self._poll_id = None

        self.root.destroy()


def main():
    """以模块方式直接运行时的入口：创建根窗口并启动工作台。"""
    root = tk.Tk()
    Workbench(root)
    root.mainloop()


if __name__ == "__main__":
    main()
