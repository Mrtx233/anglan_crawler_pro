# ======================== GUI 入口 ========================
#
# 职责: 界面构建、事件绑定、线程调度、回调转发。
#       采集业务逻辑已抽到 scraper_runner.ScraperRunner。
# 依赖: config、logger、file_utils、converter、scraper_runner。
#
# 运行: cd shopify_scraper_workbench && python main.py
#
# 说明:
#   - 原 _run_task 的业务逻辑整体移到 scraper_runner.ScraperRunner。
#   - ScraperRunner 不导入 tkinter，GUI 只负责把回调 after 转发到主线程。
#   - 新增关窗保护：_on_close 先设 stop_event，再用 root.after 轮询等待
#     采集线程结束，最后销毁窗口；加超时兜底。
# ==========================================================

import os
import sys
import threading
import time

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# 保证以脚本方式直接运行时能找到同目录模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
import converter
import file_utils
import logger
import scraper_runner

# 与原实现中裸名调用保持一致的别名
UI_COLORS = config.UI_COLORS
UI_FONT = config.UI_FONT
MONO_FONT = config.MONO_FONT

save_xlsx = file_utils.save_xlsx
build_output_folder_path = file_utils.build_output_folder_path

# 关窗等待保存的最长秒数；超时则强制退出并打日志
CLOSE_WAIT_TIMEOUT = 60
# 轮询采集线程的间隔（毫秒）
CLOSE_POLL_INTERVAL_MS = 200


class ScraperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Shopify 商品采集工作台")
        self.root.geometry("980x760")
        self.root.minsize(900, 680)
        self.root.configure(bg=UI_COLORS["background"])

        self.style = ttk.Style()
        if "clam" in self.style.theme_names():
            self.style.theme_use("clam")
        self._configure_styles()

        self.is_running = False
        self._closing = False
        self.stop_event = threading.Event()
        self.runner = None
        self._worker_thread = None
        self.pending_merge = None
        self.file_rows = []
        self.task_files = []
        self.task_folder = ""
        self.active_file_index = None

        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value="请选择分类文件夹后点击“读取文件夹”")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_percent_var = tk.StringVar(value="0%")
        self.file_count_var = tk.StringVar(value="0 个文件")

        self._build_log_window()
        logger.install_stdout_redirect(self.log_text)
        logger.install_gui_sink(self._append_log_text)
        self._build_ui()
        self._center_window()
        # 采集中直接关窗会硬杀线程、残留 Chromium 进程并丢失数据，
        # 故拦截关闭事件，先触发停止与保存。
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ────────────────────────────────────────────────────────
    # 样式与通用组件
    # ────────────────────────────────────────────────────────
    def _configure_styles(self):
        self.style.configure(
            "App.TEntry",
            fieldbackground=UI_COLORS["card_alt"],
            foreground=UI_COLORS["text"],
            bordercolor=UI_COLORS["border"],
            lightcolor=UI_COLORS["border"],
            darkcolor=UI_COLORS["border"],
            insertcolor=UI_COLORS["primary"],
            padding=(10, 9),
            font=(UI_FONT, 10),
        )
        self.style.map(
            "App.TEntry",
            bordercolor=[("focus", UI_COLORS["border_focus"])],
            lightcolor=[("focus", UI_COLORS["border_focus"])],
            darkcolor=[("focus", UI_COLORS["border_focus"])],
        )
        self.style.configure(
            "Blue.Horizontal.TProgressbar",
            troughcolor="#E8EEF7",
            background=UI_COLORS["primary"],
            bordercolor="#E8EEF7",
            lightcolor=UI_COLORS["primary"],
            darkcolor=UI_COLORS["primary"],
            thickness=10,
        )
        self.style.configure(
            "App.Vertical.TScrollbar",
            troughcolor=UI_COLORS["card"],
            background="#CBD5E1",
            bordercolor=UI_COLORS["card"],
            arrowcolor=UI_COLORS["text_secondary"],
            gripcount=0,
        )
        self.style.map(
            "App.Vertical.TScrollbar",
            background=[("active", "#94A3B8")],
        )

    def _center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = max((screen_width - width) // 2, 0)
        y = max((screen_height - height) // 2 - 20, 0)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _create_card(self, parent, **pack_kwargs):
        card = tk.Frame(
            parent,
            bg=UI_COLORS["card"],
            highlightbackground=UI_COLORS["border"],
            highlightcolor=UI_COLORS["border"],
            highlightthickness=1,
            bd=0,
        )
        card.pack(**pack_kwargs)
        return card

    def _create_button(
        self,
        parent,
        text,
        command,
        kind="secondary",
        width=None,
        padx=16,
        pady=9,
        font_size=10,
    ):
        styles = {
            "primary": {
                "bg": UI_COLORS["primary"],
                "fg": "#FFFFFF",
                "hover": UI_COLORS["primary_hover"],
                "disabled_bg": "#AFC7F5",
                "disabled_fg": "#F8FAFC",
                "border": 0,
            },
            "danger": {
                "bg": UI_COLORS["card"],
                "fg": UI_COLORS["danger"],
                "hover": UI_COLORS["danger_light"],
                "disabled_bg": UI_COLORS["card_alt"],
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
            "secondary": {
                "bg": UI_COLORS["card_alt"],
                "fg": UI_COLORS["text_secondary"],
                "hover": "#EEF2F7",
                "disabled_bg": "#F1F5F9",
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
            "ghost": {
                "bg": UI_COLORS["card"],
                "fg": UI_COLORS["primary"],
                "hover": UI_COLORS["primary_light"],
                "disabled_bg": UI_COLORS["card"],
                "disabled_fg": UI_COLORS["text_muted"],
                "border": 1,
            },
        }
        palette = styles[kind]
        button = tk.Button(
            parent,
            text=text,
            command=command,
            width=width,
            padx=padx,
            pady=pady,
            font=(UI_FONT, font_size, "bold" if kind == "primary" else "normal"),
            bg=palette["bg"],
            fg=palette["fg"],
            activebackground=palette["hover"],
            activeforeground=palette["fg"],
            disabledforeground=palette["disabled_fg"],
            relief=tk.FLAT,
            bd=0,
            highlightthickness=palette["border"],
            highlightbackground=(
                UI_COLORS["danger_border"]
                if kind == "danger"
                else UI_COLORS["border"]
            ),
            highlightcolor=(
                UI_COLORS["danger_border"]
                if kind == "danger"
                else UI_COLORS["border"]
            ),
            cursor="hand2",
        )
        button._ui_kind = kind
        button._normal_bg = palette["bg"]
        button._hover_bg = palette["hover"]
        button._disabled_bg = palette["disabled_bg"]
        button.bind("<Enter>", lambda _e, b=button: self._button_hover(b, True))
        button.bind("<Leave>", lambda _e, b=button: self._button_hover(b, False))
        return button

    @staticmethod
    def _button_hover(button, entering):
        if str(button.cget("state")) == str(tk.DISABLED):
            return
        button.configure(bg=button._hover_bg if entering else button._normal_bg)

    @staticmethod
    def _set_button_enabled(button, enabled):
        if enabled:
            button.configure(state=tk.NORMAL, bg=button._normal_bg, cursor="hand2")
        else:
            button.configure(state=tk.DISABLED, bg=button._disabled_bg, cursor="arrow")

    # ────────────────────────────────────────────────────────
    # 日志窗口
    # ────────────────────────────────────────────────────────
    def _build_log_window(self):
        self.log_win = tk.Toplevel(self.root)
        self.log_win.title("运行日志")
        self.log_win.geometry("820x480")
        self.log_win.minsize(650, 360)
        self.log_win.configure(bg=UI_COLORS["log_background"])
        self.log_win.protocol("WM_DELETE_WINDOW", self.log_win.withdraw)

        header = tk.Frame(self.log_win, bg=UI_COLORS["log_panel"], height=58)
        header.pack(fill=tk.X)
        header.pack_propagate(False)

        title_group = tk.Frame(header, bg=UI_COLORS["log_panel"])
        title_group.pack(side=tk.LEFT, padx=18, pady=10)
        tk.Label(
            title_group,
            text="运行日志",
            bg=UI_COLORS["log_panel"],
            fg="#F8FAFC",
            font=(UI_FONT, 12, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="实时查看采集、保存与转换过程",
            bg=UI_COLORS["log_panel"],
            fg="#94A3B8",
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W)

        log_actions = tk.Frame(header, bg=UI_COLORS["log_panel"])
        log_actions.pack(side=tk.RIGHT, padx=14)
        clear_btn = tk.Button(
            log_actions,
            text="清空",
            command=self._clear_logs,
            bg="#1E293B",
            fg="#CBD5E1",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=6,
            font=(UI_FONT, 9),
            cursor="hand2",
        )
        clear_btn.pack(side=tk.LEFT, padx=(0, 8))
        close_btn = tk.Button(
            log_actions,
            text="关闭",
            command=self.log_win.withdraw,
            bg="#1E293B",
            fg="#CBD5E1",
            activebackground="#334155",
            activeforeground="#FFFFFF",
            relief=tk.FLAT,
            bd=0,
            padx=14,
            pady=6,
            font=(UI_FONT, 9),
            cursor="hand2",
        )
        close_btn.pack(side=tk.LEFT)

        log_body = tk.Frame(self.log_win, bg=UI_COLORS["log_background"])
        log_body.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)

        self.log_text = scrolledtext.ScrolledText(
            log_body,
            state=tk.DISABLED,
            bg=UI_COLORS["log_background"],
            fg=UI_COLORS["log_text"],
            insertbackground="#FFFFFF",
            selectbackground="#1D4ED8",
            selectforeground="#FFFFFF",
            font=(MONO_FONT, 9),
            relief=tk.FLAT,
            bd=0,
            padx=12,
            pady=12,
            wrap=tk.WORD,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_win.withdraw()

    def _clear_logs(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def show_logs(self):
        self.log_win.deiconify()
        self.log_win.lift()
        self.log_win.focus_force()

    # ────────────────────────────────────────────────────────
    # 主界面
    # ────────────────────────────────────────────────────────
    def _build_ui(self):
        page = tk.Frame(self.root, bg=UI_COLORS["background"])
        page.pack(fill=tk.BOTH, expand=True)

        content = tk.Frame(page, bg=UI_COLORS["background"])
        content.pack(fill=tk.BOTH, expand=True, padx=24, pady=(16, 14))

        footer = tk.Frame(content, bg=UI_COLORS["background"])
        footer.pack(side=tk.BOTTOM, fill=tk.X)
        self._build_progress_card(footer)
        self._build_action_bar(footer)

        body = tk.Frame(content, bg=UI_COLORS["background"])
        body.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        self._build_header(body)
        self._build_folder_card(body)
        self._build_file_card(body)

    def _build_header(self, parent):
        header = tk.Frame(parent, bg=UI_COLORS["background"])
        header.pack(fill=tk.X, pady=(0, 12))

        brand = tk.Frame(header, bg=UI_COLORS["background"])
        brand.pack(side=tk.LEFT)

        logo = tk.Label(
            brand,
            text="S",
            width=3,
            height=1,
            bg=UI_COLORS["primary"],
            fg="#FFFFFF",
            font=(UI_FONT, 15, "bold"),
            relief=tk.FLAT,
        )
        logo.pack(side=tk.LEFT, padx=(0, 12), ipady=6)

        title_group = tk.Frame(brand, bg=UI_COLORS["background"])
        title_group.pack(side=tk.LEFT)
        tk.Label(
            title_group,
            text="Shopify 商品采集工作台",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 18, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="批量读取商品链接，自动采集、合并并转换 Shopify 数据",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 0))

        self.header_log_btn = self._create_button(
            header,
            "查看运行日志",
            self.show_logs,
            kind="ghost",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.header_log_btn.pack(side=tk.RIGHT, pady=4)

    def _build_folder_card(self, parent):
        card = self._create_card(parent, fill=tk.X, pady=(0, 12))
        inner = tk.Frame(card, bg=UI_COLORS["card"])
        inner.pack(fill=tk.X, padx=18, pady=14)

        tk.Label(
            inner,
            text="选择分类文件夹",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            inner,
            text="文件夹中需要包含待处理的 .xlsx 文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 10))

        input_row = tk.Frame(inner, bg=UI_COLORS["card"])
        input_row.pack(fill=tk.X)

        self.folder_entry = ttk.Entry(
            input_row,
            textvariable=self.folder_var,
            style="App.TEntry",
        )
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        self.browse_btn = self._create_button(
            input_row,
            "浏览文件夹",
            self._browse_folder,
            kind="secondary",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.browse_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.scan_btn = self._create_button(
            input_row,
            "读取文件夹",
            self._scan_files,
            kind="primary",
            padx=14,
            pady=8,
            font_size=9,
        )
        self.scan_btn.pack(side=tk.LEFT)

    def _build_file_card(self, parent):
        card = self._create_card(
            parent,
            fill=tk.BOTH,
            expand=True,
            pady=(0, 10),
        )
        card.configure(height=250)
        card.pack_propagate(False)
        self.file_card = card

        top = tk.Frame(card, bg=UI_COLORS["card"])
        top.pack(fill=tk.X, padx=18, pady=(10, 6))

        title_group = tk.Frame(top, bg=UI_COLORS["card"])
        title_group.pack(side=tk.LEFT)
        tk.Label(
            title_group,
            text="待处理文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
        tk.Label(
            title_group,
            text="可为每个文件单独设置需要跳过的图片位置",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(anchor=tk.W, pady=(3, 0))

        count_chip = tk.Label(
            top,
            textvariable=self.file_count_var,
            bg=UI_COLORS["primary_light"],
            fg=UI_COLORS["primary"],
            font=(UI_FONT, 9, "bold"),
            padx=10,
            pady=5,
        )
        count_chip.pack(side=tk.RIGHT)

        table_header = tk.Frame(card, bg=UI_COLORS["card_alt"], height=30)
        table_header.pack(fill=tk.X, padx=18)
        table_header.pack_propagate(False)
        table_header.grid_columnconfigure(1, weight=1)

        headers = [
            ("序号", 0, 54, tk.CENTER),
            ("文件名", 1, 0, tk.W),
            ("保留图片", 2, 100, tk.W),
            ("跳过图片", 3, 100, tk.W),
            ("状态", 4, 100, tk.CENTER),
        ]
        for text, column, width, anchor in headers:
            label = tk.Label(
                table_header,
                text=text,
                bg=UI_COLORS["card_alt"],
                fg=UI_COLORS["text_secondary"],
                font=(UI_FONT, 9, "bold"),
                anchor=anchor,
            )
            label.grid(
                row=0,
                column=column,
                sticky="nsew" if column == 1 else "ns",
                padx=(12 if column == 1 else 6),
            )
            if width:
                label.configure(width=max(width // 9, 1))

        list_shell = tk.Frame(card, bg=UI_COLORS["card"])
        list_shell.pack(fill=tk.BOTH, expand=True, padx=18, pady=(0, 8))

        self._canvas = tk.Canvas(
            list_shell,
            bg=UI_COLORS["card"],
            highlightthickness=0,
            bd=0,
        )
        scrollbar = ttk.Scrollbar(
            list_shell,
            orient=tk.VERTICAL,
            command=self._canvas.yview,
            style="App.Vertical.TScrollbar",
        )
        self.list_inner = tk.Frame(self._canvas, bg=UI_COLORS["card"])
        self._list_window = self._canvas.create_window(
            (0, 0),
            window=self.list_inner,
            anchor="nw",
        )

        self.list_inner.bind(
            "<Configure>",
            lambda _e: self._canvas.configure(scrollregion=self._canvas.bbox("all")),
        )
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfigure(self._list_window, width=e.width),
        )
        self._canvas.configure(yscrollcommand=scrollbar.set)

        self._canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._canvas.bind("<Enter>", self._bind_mousewheel)
        self._canvas.bind("<Leave>", self._unbind_mousewheel)
        self.list_inner.bind("<Enter>", self._bind_mousewheel)
        self.list_inner.bind("<Leave>", self._unbind_mousewheel)

        self._show_empty_file_state()

    def _build_progress_card(self, parent):
        card = self._create_card(parent, fill=tk.X, pady=(0, 12))
        inner = tk.Frame(card, bg=UI_COLORS["card"])
        inner.pack(fill=tk.X, padx=18, pady=14)

        progress_header = tk.Frame(inner, bg=UI_COLORS["card"])
        progress_header.pack(fill=tk.X)
        tk.Label(
            progress_header,
            text="任务进度",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 10, "bold"),
        ).pack(side=tk.LEFT)
        tk.Label(
            progress_header,
            textvariable=self.progress_percent_var,
            bg=UI_COLORS["card"],
            fg=UI_COLORS["primary"],
            font=(UI_FONT, 11, "bold"),
        ).pack(side=tk.RIGHT)

        self.progress_bar = ttk.Progressbar(
            inner,
            orient=tk.HORIZONTAL,
            mode="determinate",
            variable=self.progress_var,
            style="Blue.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill=tk.X, pady=(10, 8), ipady=1)

        status_row = tk.Frame(inner, bg=UI_COLORS["card"])
        status_row.pack(fill=tk.X)
        self.status_dot = tk.Label(
            status_row,
            text="●",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 8),
        )
        self.status_dot.pack(side=tk.LEFT, padx=(0, 7))
        self.status_label = tk.Label(
            status_row,
            textvariable=self.status_var,
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 9),
            anchor=tk.W,
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

    def _build_action_bar(self, parent):
        actions = tk.Frame(parent, bg=UI_COLORS["background"])
        actions.pack(fill=tk.X)

        hint = tk.Label(
            actions,
            text="开始后可随时打开日志窗口查看详细信息",
            bg=UI_COLORS["background"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        )
        hint.pack(side=tk.LEFT)

        button_group = tk.Frame(actions, bg=UI_COLORS["background"])
        button_group.pack(side=tk.RIGHT)

        self.log_btn = self._create_button(
            button_group,
            "查看日志",
            self.show_logs,
            kind="secondary",
            padx=16,
            pady=9,
        )
        self.log_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.stop_btn = self._create_button(
            button_group,
            "停止任务",
            self._stop,
            kind="danger",
            padx=16,
            pady=9,
        )
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 8))
        self._set_button_enabled(self.stop_btn, False)

        self.start_btn = self._create_button(
            button_group,
            "开始采集",
            self._start,
            kind="primary",
            padx=22,
            pady=9,
        )
        self.start_btn.pack(side=tk.LEFT)

        self.merge_btn = self._create_button(
            button_group,
            "合并并转换",
            self._do_merge,
            kind="primary",
            padx=22,
            pady=9,
        )
        self.merge_btn.pack(side=tk.LEFT, padx=(8, 0))
        self._set_button_enabled(self.merge_btn, False)

    # ────────────────────────────────────────────────────────
    # 文件列表交互
    # ────────────────────────────────────────────────────────
    def _bind_mousewheel(self, _event=None):
        self.root.bind_all("<MouseWheel>", self._on_mousewheel)
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux)

    def _unbind_mousewheel(self, _event=None):
        self.root.unbind_all("<MouseWheel>")
        self.root.unbind_all("<Button-4>")
        self.root.unbind_all("<Button-5>")

    def _on_mousewheel(self, event):
        if event.delta:
            self._canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _on_mousewheel_linux(self, event):
        direction = -1 if event.num == 4 else 1
        self._canvas.yview_scroll(direction, "units")

    def _show_empty_file_state(self):
        for widget in self.list_inner.winfo_children():
            widget.destroy()
        empty = tk.Frame(self.list_inner, bg=UI_COLORS["card"])
        empty.pack(fill=tk.BOTH, expand=True, pady=18)
        tk.Label(
            empty,
            text="尚未读取文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_secondary"],
            font=(UI_FONT, 10, "bold"),
        ).pack()
        tk.Label(
            empty,
            text="选择分类文件夹后，文件会显示在这里",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text_muted"],
            font=(UI_FONT, 9),
        ).pack(pady=(5, 0))

    def _browse_folder(self):
        path = filedialog.askdirectory(title="选择分类文件夹（如 lindvs.com连衣裙）")
        if path:
            self.folder_var.set(path)
            self._scan_files()

    def _scan_files(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径无效", "请先选择一个有效的分类文件夹。")
            return

        xlsx_files = sorted(
            f
            for f in os.listdir(folder)
            if f.lower().endswith(".xlsx") and not f.startswith("~")
        )

        if not xlsx_files:
            self.file_rows.clear()
            self.file_count_var.set("0 个文件")
            self._show_empty_file_state()
            self.update_status(0, 100, "该文件夹中没有找到 .xlsx 文件")
            messagebox.showinfo("未找到文件", "该文件夹下没有 .xlsx 文件。")
            return

        for widget in self.list_inner.winfo_children():
            widget.destroy()
        self.file_rows.clear()

        self.list_inner.grid_columnconfigure(0, minsize=54)
        self.list_inner.grid_columnconfigure(1, weight=1)
        self.list_inner.grid_columnconfigure(2, minsize=100)
        self.list_inner.grid_columnconfigure(3, minsize=100)
        self.list_inner.grid_columnconfigure(4, minsize=100)

        for index, filename in enumerate(xlsx_files):
            file_path = os.path.join(folder, filename)
            keep_var = tk.StringVar()
            skip_var = tk.StringVar()
            status_var = tk.StringVar(value="等待中")

            row_bg = UI_COLORS["card"] if index % 2 == 0 else "#FBFCFE"
            row = tk.Frame(
                self.list_inner,
                bg=row_bg,
                height=48,
                highlightbackground="#EEF2F7",
                highlightthickness=0,
            )
            row.grid(row=index, column=0, columnspan=5, sticky="ew")
            row.grid_columnconfigure(1, weight=1)
            row.grid_propagate(False)

            number_label = tk.Label(
                row,
                text=f"{index + 1:02d}",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(MONO_FONT, 9),
                anchor=tk.CENTER,
            )
            number_label.grid(row=0, column=0, sticky="nsew", padx=6)

            file_label = tk.Label(
                row,
                text=filename,
                bg=row_bg,
                fg=UI_COLORS["text"],
                font=(UI_FONT, 9),
                anchor=tk.W,
            )
            file_label.grid(row=0, column=1, sticky="nsew", padx=(12, 12))

            keep_shell = tk.Frame(row, bg=row_bg, width=100)
            keep_shell.grid(row=0, column=2, sticky="w", padx=(0, 6))
            keep_entry = tk.Entry(
                keep_shell,
                textvariable=keep_var,
                width=7,
                bg="#FFFFFF",
                fg=UI_COLORS["text"],
                insertbackground=UI_COLORS["primary"],
                relief=tk.FLAT,
                bd=0,
                highlightthickness=1,
                highlightbackground=UI_COLORS["primary_border"],
                highlightcolor=UI_COLORS["border_focus"],
                font=(UI_FONT, 9),
            )
            keep_entry.pack(side=tk.LEFT, ipady=5)
            tk.Label(
                keep_shell,
                text="如 1,2",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(UI_FONT, 8),
            ).pack(side=tk.LEFT, padx=(4, 0))

            skip_shell = tk.Frame(row, bg=row_bg, width=100)
            skip_shell.grid(row=0, column=3, sticky="w", padx=(0, 6))
            skip_entry = tk.Entry(
                skip_shell,
                textvariable=skip_var,
                width=7,
                bg="#FFFFFF",
                fg=UI_COLORS["text"],
                insertbackground=UI_COLORS["primary"],
                relief=tk.FLAT,
                bd=0,
                highlightthickness=1,
                highlightbackground=UI_COLORS["border"],
                highlightcolor=UI_COLORS["border_focus"],
                font=(UI_FONT, 9),
            )
            skip_entry.pack(side=tk.LEFT, ipady=5)
            tk.Label(
                skip_shell,
                text="如 1,3",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(UI_FONT, 8),
            ).pack(side=tk.LEFT, padx=(4, 0))

            status_label = tk.Label(
                row,
                textvariable=status_var,
                bg="#F1F5F9",
                fg=UI_COLORS["text_secondary"],
                font=(UI_FONT, 8, "bold"),
                padx=9,
                pady=4,
                width=7,
            )
            status_label.grid(row=0, column=4, padx=(0, 8))

            separator = tk.Frame(self.list_inner, bg="#EEF2F7", height=1)
            separator.grid(row=index, column=0, columnspan=5, sticky="sew")

            self.file_rows.append(
                {
                    "path": file_path,
                    "keep_var": keep_var,
                    "keep_entry": keep_entry,
                    "skip_var": skip_var,
                    "skip_entry": skip_entry,
                    "status_var": status_var,
                    "status_label": status_label,
                }
            )

        self.file_count_var.set(f"{len(xlsx_files)} 个文件")
        self.update_status(0, 100, f"已加载 {len(xlsx_files)} 个文件，等待开始")

    def _set_file_status(self, index, text, state="waiting"):
        if index is None or not (0 <= index < len(self.file_rows)):
            return
        row = self.file_rows[index]
        palettes = {
            "waiting": ("#F1F5F9", UI_COLORS["text_secondary"]),
            "running": (UI_COLORS["primary_light"], UI_COLORS["primary"]),
            "success": (UI_COLORS["success_light"], UI_COLORS["success"]),
            "skipped": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
            "failed": (UI_COLORS["danger_light"], UI_COLORS["danger"]),
            "stopped": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
        }
        bg, fg = palettes.get(state, palettes["waiting"])
        row["status_var"].set(text)
        row["status_label"].configure(bg=bg, fg=fg)

    # ────────────────────────────────────────────────────────
    # 任务控制
    # ────────────────────────────────────────────────────────
    def _start(self):
        if not self.file_rows:
            messagebox.showerror("无法开始", "请先选择并读取分类文件夹。")
            return

        self.task_folder = self.folder_var.get().strip()
        self.task_files = [
            {
                "path": row["path"],
                "keep": row["keep_var"].get().strip(),
                "skip": row["skip_var"].get().strip(),
            }
            for row in self.file_rows
        ]

        self.is_running = True
        self._closing = False
        self.stop_event.clear()
        self.pending_merge = None
        self._set_button_enabled(self.merge_btn, False)
        self.active_file_index = None
        self._set_running_controls(True)

        for index in range(len(self.file_rows)):
            self._set_file_status(index, "等待中", "waiting")

        self._clear_logs()
        self.update_status(0, 100, "正在初始化任务...")
        self.status_dot.configure(fg=UI_COLORS["primary"])

        self.runner = scraper_runner.ScraperRunner(
            task_files=self.task_files,
            task_folder=self.task_folder,
            stop_event=self.stop_event,
            on_file_status=self._on_file_status,
            on_progress=self._on_progress,
            on_pending_merge=self._on_pending_merge,
        )
        self._worker_thread = threading.Thread(
            target=self._worker, daemon=True, name="scraper-worker"
        )
        self._worker_thread.start()

    def _stop(self):
        if self.is_running:
            self.stop_event.set()
            self._set_button_enabled(self.stop_btn, False)
            self.status_dot.configure(fg=UI_COLORS["warning"])
            self.update_status(
                self.progress_var.get(),
                self.progress_bar["maximum"],
                "正在停止并保存当前进度...",
            )
            print("\n[系统] 用户已手动触发停止，等待当前项完成...")

    def _set_running_controls(self, running):
        self._set_button_enabled(self.start_btn, not running)
        self._set_button_enabled(self.stop_btn, running)
        self._set_button_enabled(self.browse_btn, not running)
        self._set_button_enabled(self.scan_btn, not running)
        self.folder_entry.configure(state=tk.DISABLED if running else tk.NORMAL)
        for row in self.file_rows:
            row["keep_entry"].configure(state=tk.DISABLED if running else tk.NORMAL)
            row["skip_entry"].configure(state=tk.DISABLED if running else tk.NORMAL)

    def update_status(self, current, total, text=""):
        safe_total = max(float(total or 0), 0)
        safe_current = max(float(current or 0), 0)

        if safe_total > 0:
            safe_current = min(safe_current, safe_total)
            percent = int(safe_current / safe_total * 100)
            self.progress_bar["maximum"] = safe_total
            self.progress_var.set(safe_current)
            self.progress_percent_var.set(f"{percent}%")
            self.status_var.set(f"{text}  ·  {int(safe_current)}/{int(safe_total)}")
        else:
            self.progress_bar["maximum"] = 100
            self.progress_var.set(0)
            self.progress_percent_var.set("0%")
            self.status_var.set(text)

    def _reset_buttons(self):
        self._set_running_controls(False)
        self.status_dot.configure(
            fg=UI_COLORS["warning"] if self.stop_event.is_set() else UI_COLORS["success"]
        )

    def _do_merge(self):
        if self.is_running:
            messagebox.showwarning("提示", "采集任务进行中，无法合并。")
            return
        if not self.pending_merge:
            messagebox.showwarning("提示", "暂无待合并的采集数据。")
            return
        self._set_button_enabled(self.merge_btn, False)
        threading.Thread(target=self._merge_worker, daemon=True).start()

    def _merge_worker(self):
        pending = self.pending_merge
        try:
            folder = pending["folder"]
            all_results = pending["results"]

            merge_path = os.path.join(
                build_output_folder_path(folder),
                f"{os.path.basename(folder)}_合并.xlsx",
            )
            os.makedirs(os.path.dirname(merge_path), exist_ok=True)
            save_xlsx(all_results, merge_path)
            print(f"\n{'=' * 50}")
            print(f"合并完成: {merge_path} ({len(all_results)} 条)")

            print(f"\n{'=' * 50}")
            print("开始 Shopify 转换 + 价格匹配...")
            converter.styles_to_shopify(merge_path, print)
            print("Shopify 转换 + 价格匹配完成")
            self.pending_merge = None
            self.root.after(0, self.update_status, 0, 0, "合并并转换完成")
        except Exception as exc:
            print(f"合并/转换出错: {exc}")
            import traceback
            traceback.print_exc()
            # 失败允许重试
            self.root.after(0, self._set_button_enabled, self.merge_btn, True)

    # ────────────────────────────────────────────────────────
    # 日志 sink 与 ScraperRunner 回调（均在主线程执行）
    # ────────────────────────────────────────────────────────
    def _append_log_text(self, text):
        """logger 的 GUI sink：可能由采集线程调用，统一 after 转发到主线程。"""
        if not text:
            return
        widget = getattr(self, "log_text", None)
        if widget is None:
            return
        try:
            self.root.after(0, self._write_log_text, text)
        except (tk.TclError, RuntimeError):
            pass

    def _write_log_text(self, text):
        try:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, text if text.endswith("\n") else text + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        except tk.TclError:
            pass

    def _on_file_status(self, index, text, state):
        """ScraperRunner 的文件状态回调，转发到主线程。"""
        try:
            self.root.after(0, self._set_file_status, index, text, state)
        except (tk.TclError, RuntimeError):
            pass

    def _on_progress(self, current, total, text):
        """ScraperRunner 的进度回调，转发到主线程。"""
        try:
            self.root.after(0, self.update_status, current, total, text)
        except (tk.TclError, RuntimeError):
            pass

    def _on_pending_merge(self, pending):
        """ScraperRunner 的待合并回调，转发到主线程并启用合并按钮。"""

        def _apply():
            self.pending_merge = pending
            self._set_button_enabled(self.merge_btn, True)

        try:
            self.root.after(0, _apply)
        except (tk.TclError, RuntimeError):
            pass

    # ────────────────────────────────────────────────────────
    # 工作线程与退出
    # ────────────────────────────────────────────────────────
    def _worker(self):
        error = None
        try:
            self.runner.run()
        except Exception as exc:
            error = exc
        finally:
            self.is_running = False
            if not self._closing:
                try:
                    self.root.after(0, self._finalize_worker, error)
                except (tk.TclError, RuntimeError):
                    pass

    def _finalize_worker(self, error=None):
        self._reset_buttons()
        if error is not None:
            self.status_dot.configure(fg=UI_COLORS["danger"])

    def _on_close(self):
        """关闭窗口：采集进行中先确认，再请求停止并等待保存完成后销毁。"""
        if self._closing:
            return

        if self.is_running:
            if not messagebox.askyesno(
                "确认退出",
                "采集任务正在进行中。\n"
                "退出将停止采集并保存当前文件已采集的结果，\n"
                "正在请求中的商品会被放弃。\n\n确定退出吗？",
            ):
                return

            self._closing = True
            self._set_button_enabled(self.start_btn, False)
            self._set_button_enabled(self.stop_btn, False)
            self._set_button_enabled(self.browse_btn, False)
            self._set_button_enabled(self.scan_btn, False)
            self.update_status(
                self.progress_var.get(),
                self.progress_bar["maximum"],
                "正在停止并保存，请勿关闭窗口...",
            )
            print("\n[关闭] 正在停止采集并等待保存完成...")

            self.stop_event.set()

            self._wait_worker_then_shutdown(
                deadline=time.monotonic() + CLOSE_WAIT_TIMEOUT
            )
            return

        self._shutdown()

    def _wait_worker_then_shutdown(self, deadline):
        """轮询等待采集线程结束，再销毁窗口；主线程保持响应。"""
        thread = self._worker_thread
        if thread is None or not thread.is_alive():
            self._shutdown()
            return

        if time.monotonic() > deadline:
            print(
                f"[警告] 等待保存超过 {CLOSE_WAIT_TIMEOUT} 秒，强制退出，"
                "可能丢失未落盘数据"
            )
            self._shutdown()
            return

        try:
            self.root.after(
                CLOSE_POLL_INTERVAL_MS,
                lambda: self._wait_worker_then_shutdown(deadline),
            )
        except (tk.TclError, RuntimeError):
            self._shutdown()

    def _shutdown(self):
        """注销日志 sink、还原 stdout，再销毁窗口。"""
        try:
            logger.clear_gui_sink()
        finally:
            logger.restore_stdout()
        try:
            self.root.destroy()
        except tk.TclError:
            pass


def main():
    root = tk.Tk()
    ScraperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()