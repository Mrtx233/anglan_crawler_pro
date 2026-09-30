# -*- coding: utf-8 -*-

import json
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk
from urllib.parse import urlparse


JSONL_PATH = Path(__file__).parent / "links.jsonl"


# ============================================================
# Theme
# ============================================================

C = {
    "bg": "#F6F8FB",
    "sidebar": "#0F172A",
    "sidebar_hover": "#1E293B",

    "card": "#FFFFFF",
    "soft": "#F8FAFC",
    "line": "#E5E7EB",

    "text": "#111827",
    "sub": "#475569",
    "muted": "#94A3B8",

    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "primary_soft": "#EFF6FF",

    "green": "#16A34A",
    "red": "#DC2626",
    "amber": "#D97706",
    "purple": "#7C3AED",
    "cyan": "#0891B2",

    "white": "#FFFFFF",
}


FONT = "Microsoft YaHei UI"

FONT_BODY = (FONT, 10)
FONT_SMALL = (FONT, 9)
FONT_TITLE = (FONT, 18, "bold")
FONT_SECTION = (FONT, 11, "bold")
FONT_MONO = ("Consolas", 10)


STATUS_NEW = "未使用"

STATUS_PRESETS = [
    "可使用",
    "待定",
    "量少",
    "logo",
    "404",
]


STATUS_THEME = {
    "未使用": ("#64748B", "#F1F5F9"),
    "可使用": ("#15803D", "#DCFCE7"),
    "待定": ("#0E7490", "#CFFAFE"),
    "量少": ("#B45309", "#FEF3C7"),
    "logo": ("#6D28D9", "#EDE9FE"),
    "404": ("#B91C1C", "#FEE2E2"),
    "历史，不在启用": ("#64748B", "#F1F5F9"),
}


# ============================================================
# Data
# ============================================================

def load_records():
    records = []

    if not JSONL_PATH.exists():
        return records

    for line in JSONL_PATH.read_text(encoding="utf-8").splitlines():
        if not line.strip():
            continue

        try:
            records.append(json.loads(line))
        except json.JSONDecodeError:
            continue

    return records


def build_tree(records):
    tree = {}

    for record in records:
        batch = record.get("batch", "(无批次)")
        type_name = record.get("type", "(无分类)")

        tree.setdefault(batch, {})
        tree[batch].setdefault(type_name, [])
        tree[batch][type_name].append(record)

    return tree


def normalize_link(link):
    parsed = urlparse(
        link if "//" in link else f"https://{link}"
    )

    host = (
        parsed.netloc
        or parsed.path.split("/")[0]
    ).lower()

    if host.startswith("www."):
        host = host[4:]

    return host, parsed.path.rstrip("/")


def parse_input_line(line):
    parts = line.split()

    # 兼容旧格式：id link type
    if parts and parts[0].isdigit():
        parts = parts[1:]

    if len(parts) < 2:
        return None

    link = parts[0]

    if (
        not link.startswith(("http://", "https://"))
        and "." not in link
    ):
        return None

    return link, " ".join(parts[1:])


def status_fg(status):
    return STATUS_THEME.get(
        status,
        (C["sub"], C["soft"])
    )[0]


def status_bg(status):
    return STATUS_THEME.get(
        status,
        (C["sub"], C["soft"])
    )[1]


# ============================================================
# App
# ============================================================

class LinkBrowserApp:

    def __init__(self, root):

        self.root = root

        self.root.title("链接库")

        self.root.geometry("1180x760")
        self.root.minsize(980, 620)

        self.root.configure(bg=C["bg"])

        self.records = load_records()
        self.tree = build_tree(self.records)

        self.path_stack = []
        self._current_items = []

        self.status_filter = "全部"
        self._detail_record = None

        self.current_page = "library"

        self._setup_styles()
        self._build_app()

        self.show_batches()
        self.switch_page("library")

    # ========================================================
    # ttk
    # ========================================================

    def _setup_styles(self):

        style = ttk.Style()

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Modern.Treeview",
            background=C["card"],
            foreground=C["text"],
            fieldbackground=C["card"],
            borderwidth=0,
            rowheight=38,
            font=FONT_BODY,
        )

        style.configure(
            "Modern.Treeview.Heading",
            background=C["soft"],
            foreground=C["sub"],
            borderwidth=0,
            relief="flat",
            font=(FONT, 9, "bold"),
            padding=(10, 9),
        )

        style.map(
            "Modern.Treeview",
            background=[
                ("selected", C["primary_soft"])
            ],
            foreground=[
                ("selected", C["primary"])
            ],
        )

        style.configure(
            "Modern.TEntry",
            fieldbackground=C["white"],
            foreground=C["text"],
            bordercolor=C["line"],
            lightcolor=C["line"],
            darkcolor=C["line"],
            padding=7,
        )

        style.configure(
            "Modern.TCombobox",
            fieldbackground=C["white"],
            background=C["white"],
            foreground=C["text"],
            bordercolor=C["line"],
            lightcolor=C["line"],
            darkcolor=C["line"],
            padding=6,
        )

        style.map(
            "Modern.TCombobox",
            fieldbackground=[
                ("readonly", C["white"])
            ],
            foreground=[
                ("readonly", C["text"])
            ],
        )

    # ========================================================
    # Common
    # ========================================================

    def button(
        self,
        parent,
        text,
        command,
        kind="secondary",
        padx=14,
        pady=7,
        font=None,
    ):

        if kind == "primary":
            bg = C["primary"]
            fg = C["white"]
            active = C["primary_hover"]
            highlight = 0
        else:
            bg = C["white"]
            fg = C["sub"]
            active = C["soft"]
            highlight = 1

        return tk.Button(
            parent,
            text=text,
            command=command,

            bg=bg,
            fg=fg,

            activebackground=active,
            activeforeground=fg,

            relief=tk.FLAT,
            bd=0,

            highlightthickness=highlight,
            highlightbackground=C["line"],

            font=font or FONT_BODY,

            padx=padx,
            pady=pady,

            cursor="hand2",
        )

    def card(self, parent):

        return tk.Frame(
            parent,
            bg=C["card"],
            highlightthickness=1,
            highlightbackground=C["line"],
        )

    # ========================================================
    # Shell
    # ========================================================

    def _build_app(self):

        self.sidebar = tk.Frame(
            self.root,
            bg=C["sidebar"],
            width=190,
        )

        self.sidebar.pack(
            side=tk.LEFT,
            fill=tk.Y,
        )

        self.sidebar.pack_propagate(False)

        self.main = tk.Frame(
            self.root,
            bg=C["bg"],
        )

        self.main.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True,
        )

        self._build_sidebar()
        self._build_header()

        self.page_host = tk.Frame(
            self.main,
            bg=C["bg"],
        )

        self.page_host.pack(
            fill=tk.BOTH,
            expand=True,
            padx=24,
            pady=(0, 8),
        )

        self.library_page = tk.Frame(
            self.page_host,
            bg=C["bg"],
        )

        self.add_page = tk.Frame(
            self.page_host,
            bg=C["bg"],
        )

        self._build_library_page()
        self._build_add_page()

        self._build_statusbar()

    # ========================================================
    # Sidebar
    # ========================================================

    def _build_sidebar(self):

        brand = tk.Frame(
            self.sidebar,
            bg=C["sidebar"],
        )

        brand.pack(
            fill=tk.X,
            padx=18,
            pady=(22, 28),
        )

        logo = tk.Label(
            brand,
            text="L",
            width=2,

            bg=C["primary"],
            fg=C["white"],

            font=(FONT, 15, "bold"),
        )

        logo.pack(side=tk.LEFT)

        tk.Label(
            brand,
            text="链接库",

            bg=C["sidebar"],
            fg=C["white"],

            font=(FONT, 12, "bold"),
        ).pack(
            side=tk.LEFT,
            padx=(10, 0),
        )

        self.nav_buttons = {}

        self.nav_buttons["library"] = self._create_nav(
            "library",
            "▦",
            "链接库",
        )

        self.nav_buttons["add"] = self._create_nav(
            "add",
            "+",
            "新增链接",
        )

        tk.Frame(
            self.sidebar,
            bg=C["sidebar"],
        ).pack(
            fill=tk.BOTH,
            expand=True,
        )

        footer = tk.Frame(
            self.sidebar,
            bg=C["sidebar"],
        )

        footer.pack(
            fill=tk.X,
            padx=18,
            pady=18,
        )

        self.sidebar_count = tk.Label(
            footer,
            text="",

            bg=C["sidebar"],
            fg="#64748B",

            font=(FONT, 8),
        )

        self.sidebar_count.pack(
            anchor="w"
        )

        self._update_sidebar_count()

    def _create_nav(
        self,
        key,
        icon,
        title,
    ):

        box = tk.Frame(
            self.sidebar,
            bg=C["sidebar"],
            cursor="hand2",
        )

        box.pack(
            fill=tk.X,
            padx=10,
            pady=3,
        )

        icon_label = tk.Label(
            box,
            text=icon,
            width=3,

            bg=C["sidebar"],
            fg="#94A3B8",

            font=(FONT, 12, "bold"),

            cursor="hand2",
        )

        icon_label.pack(
            side=tk.LEFT,
            padx=(8, 0),
            pady=10,
        )

        title_label = tk.Label(
            box,
            text=title,

            bg=C["sidebar"],
            fg="#CBD5E1",

            font=FONT_BODY,

            cursor="hand2",
        )

        title_label.pack(
            side=tk.LEFT,
            padx=(4, 0),
        )

        for widget in (
            box,
            icon_label,
            title_label,
        ):
            widget.bind(
                "<Button-1>",
                lambda e, page=key:
                self.switch_page(page),
            )

        box.icon_label = icon_label
        box.title_label = title_label

        return box

    def _set_active_nav(self, key):

        for nav_key, box in self.nav_buttons.items():

            active = nav_key == key

            bg = (
                C["primary"]
                if active
                else C["sidebar"]
            )

            box.config(bg=bg)

            box.icon_label.config(
                bg=bg,
                fg=(
                    C["white"]
                    if active
                    else "#94A3B8"
                ),
            )

            box.title_label.config(
                bg=bg,
                fg=(
                    C["white"]
                    if active
                    else "#CBD5E1"
                ),
                font=(
                    FONT,
                    10,
                    "bold" if active else "normal",
                ),
            )

    # ========================================================
    # Header
    # ========================================================

    def _build_header(self):

        header = tk.Frame(
            self.main,
            bg=C["bg"],
            height=78,
        )

        header.pack(
            fill=tk.X,
            padx=24,
            pady=(18, 8),
        )

        header.pack_propagate(False)

        self.page_title = tk.Label(
            header,
            text="链接库",

            bg=C["bg"],
            fg=C["text"],

            font=FONT_TITLE,
        )

        self.page_title.pack(
            side=tk.LEFT,
            anchor="w",
            pady=8,
        )

    # ========================================================
    # Statusbar
    # ========================================================

    def _build_statusbar(self):

        bar = tk.Frame(
            self.main,
            bg=C["bg"],
        )

        bar.pack(
            side=tk.BOTTOM,
            fill=tk.X,
            padx=26,
            pady=(0, 10),
            before=self.page_host,
        )

        self.statusbar = tk.Label(
            bar,
            text="就绪",

            bg=C["bg"],
            fg=C["muted"],

            font=(FONT, 8),
        )

        self.statusbar.pack(
            side=tk.LEFT
        )

    # ========================================================
    # Switch page
    # ========================================================

    def switch_page(self, page):

        self.current_page = page
        self._set_active_nav(page)

        self.library_page.pack_forget()
        self.add_page.pack_forget()

        if page == "library":

            self.page_title.config(
                text="链接库"
            )

            self.library_page.pack(
                fill=tk.BOTH,
                expand=True,
            )

        else:

            self.page_title.config(
                text="新增链接"
            )

            self.add_page.pack(
                fill=tk.BOTH,
                expand=True,
            )

    # ========================================================
    # Library page
    # ========================================================

    def _build_library_page(self):

        workspace = tk.Frame(
            self.library_page,
            bg=C["bg"],
        )

        workspace.pack(
            fill=tk.BOTH,
            expand=True,
        )

        table_card = self.card(
            workspace
        )

        table_card.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True,
        )

        detail_card = self.card(
            workspace
        )

        detail_card.pack(
            side=tk.RIGHT,
            fill=tk.Y,
            padx=(14, 0),
            before=table_card,
        )

        detail_card.config(
            width=220
        )

        detail_card.pack_propagate(False)

        self._build_table(table_card)
        self._build_detail(detail_card)

    # ========================================================
    # Table
    # ========================================================

    def _build_table(self, parent):

        # Topbar
        top = tk.Frame(
            parent,
            bg=C["card"],
        )

        top.pack(
            fill=tk.X,
            padx=16,
            pady=(14, 10),
        )

        self.back_btn = self.button(
            top,
            "←",
            self.go_back,
            padx=10,
            pady=5,
        )

        self.back_btn.pack(
            side=tk.LEFT
        )

        self.path_label = tk.Label(
            top,
            text="全部批次",

            bg=C["card"],
            fg=C["text"],

            font=FONT_SECTION,
        )

        self.path_label.pack(
            side=tk.LEFT,
            padx=(12, 0),
        )

        self.path_meta = tk.Label(
            top,
            text="",

            bg=C["card"],
            fg=C["muted"],

            font=(FONT, 8),
        )

        self.path_meta.pack(
            side=tk.LEFT,
            padx=(8, 0),
        )

        # Search
        tools = tk.Frame(
            parent,
            bg=C["card"],
        )

        tools.pack(
            fill=tk.X,
            padx=16,
            pady=(0, 10),
        )

        search_box = tk.Frame(
            tools,
            bg=C["white"],
            highlightthickness=1,
            highlightbackground=C["line"],
        )

        search_box.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
        )

        tk.Label(
            search_box,
            text="⌕",

            bg=C["white"],
            fg=C["muted"],

            font=(FONT, 13),
        ).pack(
            side=tk.LEFT,
            padx=(10, 5),
        )

        self.search_var = tk.StringVar()

        self.search_entry = tk.Entry(
            search_box,

            textvariable=self.search_var,

            bg=C["white"],
            fg=C["text"],

            insertbackground=C["text"],

            relief=tk.FLAT,
            bd=0,

            font=FONT_BODY,
        )

        self.search_entry.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
            ipady=8,
        )

        self.search_entry.bind(
            "<KeyRelease>",
            lambda e: self.on_search_changed(),
        )

        self.filter_combo = ttk.Combobox(
            tools,

            style="Modern.TCombobox",

            state="readonly",

            width=11,

            values=["全部"],

            font=FONT_SMALL,
        )

        self.filter_combo.set("全部")

        self.filter_combo.pack(
            side=tk.LEFT,
            padx=(10, 0),
        )

        self.filter_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self.on_filter_changed(),
        )

        # Treeview
        tree_frame = tk.Frame(
            parent,

            bg=C["card"],

            highlightthickness=1,
            highlightbackground=C["line"],
        )

        tree_frame.pack(
            fill=tk.BOTH,
            expand=True,
            padx=16,
            pady=(0, 12),
        )

        columns = (
            "id",
            "link",
            "status",
            "note",
            "action",
        )

        self.listbox = ttk.Treeview(
            tree_frame,

            columns=columns,

            show="headings",

            selectmode="extended",

            style="Modern.Treeview",
        )

        self.listbox.column(
            "id",
            anchor="center",
        )

        self.listbox.column(
            "link",
            anchor="w",
        )

        self.listbox.column(
            "status",
            anchor="center",
        )

        self.listbox.column(
            "note",
            anchor="w",
        )

        self.listbox.column(
            "action",
            anchor="center",
        )

        scrollbar = ttk.Scrollbar(
            tree_frame,
            orient="vertical",
            command=self.listbox.yview,
        )

        self.listbox.configure(
            yscrollcommand=scrollbar.set
        )

        self.listbox.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True,
        )

        scrollbar.pack(
            side=tk.RIGHT,
            fill=tk.Y,
        )

        self.listbox.bind(
            "<Button-1>",
            self.on_click,
        )

        self.listbox.bind(
            "<Double-Button-1>",
            self.on_double_click,
        )

        self.listbox.bind(
            "<Return>",
            self.on_double_click,
        )

        self.listbox.bind(
            "<<TreeviewSelect>>",
            self.on_select,
        )

        self.listbox.bind(
            "<Configure>",
            self._resize_columns,
        )

        # Batch action
        actions = tk.Frame(
            parent,
            bg=C["soft"],
        )

        actions.pack(
            side=tk.BOTTOM,
            fill=tk.X,
            padx=16,
            pady=(0, 14),
            before=tree_frame,
        )

        top_action = tk.Frame(
            actions,
            bg=C["soft"],
        )

        top_action.pack(
            fill=tk.X,
            padx=12,
            pady=(8, 6),
        )

        tk.Label(
            top_action,
            text="批量操作",

            bg=C["soft"],
            fg=C["sub"],

            font=(FONT, 9, "bold"),
        ).pack(
            side=tk.LEFT
        )

        self.selection_count = tk.Label(
            top_action,
            text="未选择",

            bg=C["soft"],
            fg=C["muted"],

            font=(FONT, 8),
        )

        self.selection_count.pack(
            side=tk.LEFT,
            padx=(8, 0),
        )

        # Status buttons
        status_row = tk.Frame(
            actions,
            bg=C["soft"],
        )

        status_row.pack(
            fill=tk.X,
            padx=12,
            pady=(0, 6),
        )

        for status in STATUS_PRESETS:

            tk.Button(
                status_row,

                text=status,

                command=lambda s=status:
                self.set_selected_status(s),

                bg=status_bg(status),
                fg=status_fg(status),

                activebackground=status_bg(status),
                activeforeground=status_fg(status),

                relief=tk.FLAT,
                bd=0,

                font=(FONT, 8, "bold"),

                padx=9,
                pady=4,

                cursor="hand2",
            ).pack(
                side=tk.LEFT,
                padx=(0, 5),
            )

        self.custom_status_entry = ttk.Entry(
            status_row,

            style="Modern.TEntry",

            width=11,

            font=FONT_SMALL,
        )

        self.custom_status_entry.pack(
            side=tk.LEFT,
            padx=(4, 4),
        )

        self.custom_status_entry.bind(
            "<Return>",
            lambda e: self.set_custom_status(),
        )

        self.button(
            status_row,
            "应用",
            self.set_custom_status,
            kind="primary",
            padx=10,
            pady=5,
            font=FONT_SMALL,
        ).pack(
            side=tk.LEFT
        )

        # Note
        note_row = tk.Frame(
            actions,
            bg=C["soft"],
        )

        note_row.pack(
            fill=tk.X,
            padx=12,
            pady=(0, 9),
        )

        self.note_entry = ttk.Entry(
            note_row,

            style="Modern.TEntry",

            font=FONT_SMALL,
        )

        self.note_entry.pack(
            side=tk.LEFT,
            fill=tk.X,
            expand=True,
        )

        self.note_entry.bind(
            "<Return>",
            lambda e: self.apply_note(),
        )

        self.button(
            note_row,
            "保存备注",
            self.apply_note,
            padx=10,
            pady=5,
            font=FONT_SMALL,
        ).pack(
            side=tk.LEFT,
            padx=(6, 12),
        )

        self.approver_entry = ttk.Entry(
            note_row,

            style="Modern.TEntry",

            width=12,

            font=FONT_SMALL,
        )

        self.approver_entry.pack(
            side=tk.LEFT
        )

        self.button(
            note_row,
            "保存审批人",
            self.apply_approver,
            padx=10,
            pady=5,
            font=FONT_SMALL,
        ).pack(
            side=tk.LEFT,
            padx=(6, 0),
        )

    # ========================================================
    # Detail
    # ========================================================

    def _build_detail(self, parent):

        tk.Label(
            parent,
            text="详情",

            bg=C["card"],
            fg=C["text"],

            font=FONT_SECTION,
        ).pack(
            anchor="w",
            padx=18,
            pady=(16, 12),
        )

        tk.Frame(
            parent,
            bg=C["line"],
            height=1,
        ).pack(
            fill=tk.X
        )

        self.detail_host = tk.Frame(
            parent,
            bg=C["card"],
        )

        self.detail_host.pack(
            fill=tk.BOTH,
            expand=True,
            padx=18,
            pady=14,
        )

        self.detail_canvas = tk.Canvas(
            self.detail_host,
            bg=C["card"],
            highlightthickness=0,
            bd=0,
        )

        detail_scrollbar = ttk.Scrollbar(
            self.detail_host,
            orient="vertical",
            command=self.detail_canvas.yview,
        )

        self.detail_canvas.configure(
            yscrollcommand=detail_scrollbar.set,
        )

        self.detail_canvas.pack(
            side=tk.LEFT,
            fill=tk.BOTH,
            expand=True,
        )

        detail_scrollbar.pack(
            side=tk.RIGHT,
            fill=tk.Y,
        )

        self.detail_inner = tk.Frame(
            self.detail_canvas,
            bg=C["card"],
        )

        self._detail_window = self.detail_canvas.create_window(
            (0, 0),
            window=self.detail_inner,
            anchor="nw",
        )

        self.detail_inner.bind(
            "<Configure>",
            lambda e: self.detail_canvas.configure(
                scrollregion=self.detail_canvas.bbox("all"),
            ),
        )

        self.detail_canvas.bind(
            "<Configure>",
            lambda e: self.detail_canvas.itemconfigure(
                self._detail_window,
                width=e.width,
            ),
        )

        # 鼠标滚轮
        def _on_detail_mousewheel(event):
            if event.delta:
                self.detail_canvas.yview_scroll(
                    int(-event.delta / 120), "units",
                )

        self.detail_canvas.bind(
            "<Enter>",
            lambda e: self.detail_canvas.bind_all(
                "<MouseWheel>", _on_detail_mousewheel,
            ),
        )

        self.detail_canvas.bind(
            "<Leave>",
            lambda e: self.detail_canvas.unbind_all(
                "<MouseWheel>",
            ),
        )

        self.detail_empty = tk.Frame(
            self.detail_inner,
            bg=C["card"],
        )

        self.detail_empty.pack(
            fill=tk.BOTH,
            expand=True,
        )

        tk.Label(
            self.detail_empty,
            text="未选择",

            bg=C["card"],
            fg=C["muted"],

            font=FONT_SMALL,
        ).pack(
            pady=100
        )

        self.detail_content = tk.Frame(
            self.detail_inner,
            bg=C["card"],
        )

        self.detail_labels = {}

        fields = [
            ("id", "ID"),
            ("batch", "批次"),
            ("type", "分类"),
            ("link", "链接"),
            ("status", "状态"),
            ("note", "备注"),
            ("approver", "审批人"),
            ("created_at", "录入时间"),
            ("updated_at", "修改时间"),
        ]

        for key, title in fields:

            box = tk.Frame(
                self.detail_content,
                bg=C["card"],
            )

            box.pack(
                fill=tk.X,
                pady=(0, 11),
            )

            tk.Label(
                box,
                text=title,

                bg=C["card"],
                fg=C["muted"],

                font=(FONT, 8),
            ).pack(
                anchor="w"
            )

            label = tk.Label(
                box,
                text="-",

                bg=C["card"],
                fg=C["sub"],

                font=FONT_SMALL,

                anchor="w",
                justify="left",

                wraplength=180,
            )

            label.pack(
                fill=tk.X,
                pady=(2, 0),
            )

            self.detail_labels[key] = label

        self.button(
            self.detail_content,
            "复制链接",
            self.copy_current_detail_link,
            kind="primary",
            pady=7,
        ).pack(
            fill=tk.X,
            pady=(5, 0),
        )

    # ========================================================
    # Add page
    # ========================================================

    def _build_add_page(self):

        card = self.card(
            self.add_page
        )

        card.pack(
            fill=tk.BOTH,
            expand=True,
        )

        form = tk.Frame(
            card,
            bg=C["card"],
        )

        form.pack(
            fill=tk.BOTH,
            expand=True,
            padx=24,
            pady=20,
        )

        # Batch
        batch_row = tk.Frame(
            form,
            bg=C["card"],
        )

        batch_row.pack(
            fill=tk.X,
            pady=(0, 14),
        )

        tk.Label(
            batch_row,
            text="批次",

            width=8,
            anchor="w",

            bg=C["card"],
            fg=C["sub"],

            font=(FONT, 9, "bold"),
        ).pack(
            side=tk.LEFT
        )

        self.batch_entry = ttk.Entry(
            batch_row,

            style="Modern.TEntry",

            width=18,

            font=FONT_BODY,
        )

        self.batch_entry.pack(
            side=tk.LEFT
        )

        if self.tree:
            self.batch_entry.insert(
                0,
                sorted(self.tree)[-1],
            )

        # Text
        input_header = tk.Frame(
            form,
            bg=C["card"],
        )

        input_header.pack(
            fill=tk.X,
            pady=(0, 6),
        )

        tk.Label(
            input_header,
            text="链接",

            bg=C["card"],
            fg=C["sub"],

            font=(FONT, 9, "bold"),
        ).pack(
            side=tk.LEFT
        )

        self.entry_counter = tk.Label(
            input_header,
            text="0 条",

            bg=C["card"],
            fg=C["muted"],

            font=(FONT, 8),
        )

        self.entry_counter.pack(
            side=tk.RIGHT
        )

        editor = tk.Frame(
            form,

            bg=C["white"],

            highlightthickness=1,
            highlightbackground=C["line"],
        )

        editor.pack(
            fill=tk.BOTH,
            expand=True,
        )

        self.entry_text = tk.Text(
            editor,

            font=FONT_MONO,

            bg=C["white"],
            fg=C["text"],

            insertbackground=C["primary"],
            selectbackground=C["primary_soft"],

            relief=tk.FLAT,
            bd=0,

            padx=14,
            pady=12,

            undo=True,
        )

        self.entry_text.pack(
            fill=tk.BOTH,
            expand=True,
        )

        self.entry_text.bind(
            "<KeyRelease>",
            lambda e: self._update_entry_count(),
        )

        # Bottom actions
        bottom = tk.Frame(
            form,
            bg=C["card"],
        )

        bottom.pack(
            side=tk.BOTTOM,
            fill=tk.X,
            pady=(14, 0),
            before=editor,
        )

        tk.Label(
            bottom,
            text="格式：链接 分类",

            bg=C["card"],
            fg=C["muted"],

            font=(FONT, 8),
        ).pack(
            side=tk.LEFT
        )

        self.button(
            bottom,
            "清空",
            self.clear_entry,
            padx=16,
            pady=7,
        ).pack(
            side=tk.RIGHT
        )

        self.button(
            bottom,
            "保存",
            self.submit_entries,
            kind="primary",
            padx=22,
            pady=7,
            font=(FONT, 10, "bold"),
        ).pack(
            side=tk.RIGHT,
            padx=(0, 8),
        )

    # ========================================================
    # Add link
    # ========================================================

    def clear_entry(self):

        self.entry_text.delete(
            "1.0",
            tk.END,
        )

        self._update_entry_count()

    def _update_entry_count(self):

        lines = [
            line
            for line in self.entry_text.get(
                "1.0",
                tk.END,
            ).splitlines()
            if line.strip()
        ]

        self.entry_counter.config(
            text=f"{len(lines)} 条"
        )

    def submit_entries(self):

        batch = self.batch_entry.get().strip()

        if not batch:
            self.set_status("请输入批次")
            return

        lines = [
            line.strip()
            for line in self.entry_text.get(
                "1.0",
                tk.END,
            ).splitlines()
            if line.strip()
        ]

        if not lines:
            self.set_status("请输入链接")
            return

        existing = {
            normalize_link(
                record.get("link", "")
            )
            for record in self.records
        }

        next_num = 1

        for record in self.records:

            try:
                next_num = max(
                    next_num,
                    int(record.get("id", 0)) + 1,
                )
            except (TypeError, ValueError):
                continue

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        added = []
        skipped = []
        invalid = []

        for line in lines:

            parsed = parse_input_line(line)

            if not parsed:
                invalid.append(line)
                continue

            link, link_type = parsed

            key = normalize_link(link)

            if key in existing:
                skipped.append(link)
                continue

            existing.add(key)

            added.append({
                "id": str(next_num).zfill(5),
                "batch": batch,
                "type": link_type,
                "link": link,
                "status": STATUS_NEW,
                "created_at": now,
                "updated_at": now,
                "note": "",
                "approver": "",
            })

            next_num += 1

        if added:

            self.records.extend(added)

            self.save_all()

            self.tree = build_tree(
                self.records
            )

            self._update_sidebar_count()

            self.clear_entry()

            self.show_batches()

        parts = []

        if added:
            parts.append(
                f"新增 {len(added)}"
            )

        if skipped:
            parts.append(
                f"重复 {len(skipped)}"
            )

        if invalid:
            parts.append(
                f"无效 {len(invalid)}"
            )

        self.set_status(
            " · ".join(parts)
        )

    # ========================================================
    # Batch actions
    # ========================================================

    def _selected_records(self):

        if len(self.path_stack) < 2:
            return []

        records = []

        for item in self.listbox.selection():

            index = self.listbox.index(
                item
            )

            if (
                0
                <= index
                < len(self._current_items)
            ):
                records.append(
                    self._current_items[index]
                )

        return records

    def set_selected_status(self, status):

        records = self._selected_records()

        if not records:
            self.set_status("请选择链接")
            return

        ids = {
            record.get("id")
            for record in records
        }

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        changed = 0

        for record in self.records:

            if record.get("id") in ids:

                if record.get("status") != status:
                    record["status"] = status
                    changed += 1

                record["updated_at"] = now

        self.save_all()

        self.tree = build_tree(
            self.records
        )

        self._refresh_links()

        self.set_status(
            f"已修改 {changed} 条"
        )

    def set_custom_status(self):

        status = (
            self.custom_status_entry
            .get()
            .strip()
        )

        if not status:
            return

        self.set_selected_status(status)

    def apply_note(self):

        self._apply_field(
            "note",
            self.note_entry.get().strip(),
            "备注",
        )

    def apply_approver(self):

        self._apply_field(
            "approver",
            self.approver_entry.get().strip(),
            "审批人",
        )

    def _apply_field(
        self,
        field,
        value,
        title,
    ):

        records = self._selected_records()

        if not records:
            self.set_status("请选择链接")
            return

        ids = {
            record.get("id")
            for record in records
        }

        now = datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        for record in self.records:

            if record.get("id") in ids:

                record[field] = value
                record["updated_at"] = now

        self.save_all()

        self.tree = build_tree(
            self.records
        )

        self._refresh_links()

        self.set_status(
            f"{title}已保存"
        )

    # ========================================================
    # Navigation
    # ========================================================

    def show_batches(self):

        self.listbox.delete(
            *self.listbox.get_children()
        )

        self.path_stack = []
        self._current_items = []

        self.path_label.config(
            text="全部批次"
        )

        self.path_meta.config(
            text=f"{len(self.tree)} 个"
        )

        self.back_btn.config(
            state=tk.DISABLED
        )

        self.search_var.set("")

        self.status_filter = "全部"

        self.filter_combo.configure(
            values=["全部"]
        )

        self.filter_combo.set("全部")

        self._set_headings(
            "批次",
            "分类",
            "链接",
        )

        self._clear_detail()

        for batch in sorted(
            self.tree,
            reverse=True,
        ):

            total = sum(
                len(items)
                for items in self.tree[
                    batch
                ].values()
            )

            self.listbox.insert(
                "",
                tk.END,

                values=(
                    batch,
                    len(self.tree[batch]),
                    total,
                    "",
                    "",
                ),
            )

        self.set_status(
            f"{len(self.records)} 条链接"
        )

    def show_types(self, batch):

        self.listbox.delete(
            *self.listbox.get_children()
        )

        self.path_stack = [
            ("batch", batch)
        ]

        self._current_items = []

        self.path_label.config(
            text=batch
        )

        self.path_meta.config(
            text=""
        )

        self.back_btn.config(
            state=tk.NORMAL
        )

        self.search_var.set("")

        self.filter_combo.configure(
            values=["全部"]
        )

        self.filter_combo.set("全部")

        self._set_headings(
            "分类",
            "链接",
            "可使用",
        )

        self._clear_detail()

        types = self.tree.get(
            batch,
            {},
        )

        ordered = sorted(
            types,
            key=lambda key: (
                -len(types[key]),
                key,
            ),
        )

        for type_name in ordered:

            items = types[
                type_name
            ]

            available = sum(
                record.get("status")
                == "可使用"
                for record in items
            )

            self.listbox.insert(
                "",
                tk.END,

                values=(
                    type_name,
                    len(items),
                    available,
                    "",
                    "",
                ),
            )

    def show_links(
        self,
        batch,
        type_name,
    ):

        self.listbox.delete(
            *self.listbox.get_children()
        )

        self.path_stack = [
            ("batch", batch),
            ("type", type_name),
        ]

        self.path_label.config(
            text=type_name
        )

        self.back_btn.config(
            state=tk.NORMAL
        )

        self._set_headings(
            "ID",
            "链接",
            "状态",
            "备注",
            "操作",
        )

        all_items = (
            self.tree
            .get(batch, {})
            .get(type_name, [])
        )

        statuses = sorted({
            record.get("status", "")
            for record in all_items
            if record.get("status")
        })

        options = [
            "全部"
        ] + statuses

        self.filter_combo.configure(
            values=options
        )

        if self.status_filter not in options:

            self.status_filter = "全部"
            self.filter_combo.set("全部")

        query = (
            self.search_var
            .get()
            .strip()
            .lower()
        )

        items = []

        for record in all_items:

            if (
                self.status_filter != "全部"
                and record.get("status")
                != self.status_filter
            ):
                continue

            text = " ".join(
                str(record.get(key, ""))
                for key in (
                    "id",
                    "link",
                    "status",
                    "note",
                    "approver",
                )
            ).lower()

            if (
                query
                and query not in text
            ):
                continue

            items.append(record)

        self._current_items = items

        for record in items:

            status = record.get(
                "status",
                "",
            )

            tag = self._status_tag(
                status
            )

            self.listbox.tag_configure(
                tag,
                foreground=status_fg(status),
            )

            self.listbox.insert(
                "",
                tk.END,

                values=(
                    record.get("id", ""),
                    record.get("link", ""),
                    status,
                    record.get("note", ""),
                    "复制",
                ),

                tags=(tag,),
            )

        self.path_meta.config(
            text=f"{len(items)}/{len(all_items)}"
        )

        self.selection_count.config(
            text="未选择"
        )

        self._clear_detail()

        self.root.after_idle(
            self._resize_columns
        )

    def _refresh_links(self):

        if len(self.path_stack) == 2:

            batch = self.path_stack[0][1]
            type_name = self.path_stack[1][1]

            self.show_links(
                batch,
                type_name,
            )

    def go_back(self):

        if len(self.path_stack) == 2:

            self.show_types(
                self.path_stack[0][1]
            )

        elif len(self.path_stack) == 1:

            self.show_batches()

    # ========================================================
    # Search
    # ========================================================

    def on_search_changed(self):

        if len(self.path_stack) == 2:
            self._refresh_links()

    def on_filter_changed(self):

        if len(self.path_stack) != 2:
            return

        self.status_filter = (
            self.filter_combo.get()
        )

        self._refresh_links()

    # ========================================================
    # Tree headings
    # ========================================================

    def _set_headings(
        self,
        c1,
        c2,
        c3,
        c4="",
        c5="",
    ):

        values = (
            c1,
            c2,
            c3,
            c4,
            c5,
        )

        columns = (
            "id",
            "link",
            "status",
            "note",
            "action",
        )

        for column, value in zip(
            columns,
            values,
        ):

            self.listbox.heading(
                column,
                text=value,
            )

        self.root.after_idle(
            self._resize_columns
        )

    # ========================================================
    # Resize
    # ========================================================

    def _resize_columns(
        self,
        event=None,
    ):

        total = self.listbox.winfo_width()

        if total < 200:
            return

        # Batch / Type
        if len(self.path_stack) < 2:

            self.listbox.column(
                "id",
                width=int(total * 0.45),
                minwidth=120,
                stretch=True,
            )

            self.listbox.column(
                "link",
                width=int(total * 0.28),
                minwidth=80,
                stretch=True,
            )

            self.listbox.column(
                "status",
                width=int(total * 0.20),
                minwidth=80,
                stretch=True,
            )

            self.listbox.column(
                "note",
                width=1,
                minwidth=1,
                stretch=False,
            )

            self.listbox.column(
                "action",
                width=1,
                minwidth=1,
                stretch=False,
            )

            return

        self.listbox.column(
            "id",
            width=80,
            minwidth=70,
            stretch=False,
        )

        self.listbox.column(
            "status",
            width=90,
            minwidth=80,
            stretch=False,
        )

        self.listbox.column(
            "action",
            width=65,
            minwidth=60,
            stretch=False,
        )

        remain = total - 235

        self.listbox.column(
            "link",
            width=int(remain * 0.68),
            minwidth=220,
            stretch=True,
        )

        self.listbox.column(
            "note",
            width=int(remain * 0.32),
            minwidth=100,
            stretch=True,
        )

    # ========================================================
    # Select
    # ========================================================

    def on_select(self, event=None):

        selected = self.listbox.selection()

        if len(self.path_stack) < 2:
            return

        if not selected:

            self.selection_count.config(
                text="未选择"
            )

            return

        self.selection_count.config(
            text=f"{len(selected)} 条"
        )

        index = self.listbox.index(
            selected[0]
        )

        if index >= len(self._current_items):
            return

        record = self._current_items[index]

        self.show_record_detail(
            record
        )

        if len(selected) == 1:

            self.note_entry.delete(
                0,
                tk.END,
            )

            self.note_entry.insert(
                0,
                record.get("note", ""),
            )

            self.approver_entry.delete(
                0,
                tk.END,
            )

            self.approver_entry.insert(
                0,
                record.get("approver", ""),
            )

    # ========================================================
    # Double click
    # ========================================================

    def on_double_click(
        self,
        event=None,
    ):

        selected = self.listbox.selection()

        if not selected:
            return

        index = self.listbox.index(
            selected[0]
        )

        if not self.path_stack:

            batches = sorted(
                self.tree,
                reverse=True,
            )

            if index < len(batches):

                self.show_types(
                    batches[index]
                )

            return

        if len(self.path_stack) == 1:

            batch = self.path_stack[0][1]

            types = self.tree.get(
                batch,
                {},
            )

            ordered = sorted(
                types,
                key=lambda key: (
                    -len(types[key]),
                    key,
                ),
            )

            if index < len(ordered):

                self.show_links(
                    batch,
                    ordered[index],
                )

            return

        if index < len(
            self._current_items
        ):

            self.show_record_detail(
                self._current_items[index]
            )

    # ========================================================
    # Click Copy
    # ========================================================

    def on_click(
        self,
        event,
    ):

        if len(self.path_stack) < 2:
            return

        if (
            self.listbox.identify_region(
                event.x,
                event.y,
            )
            != "cell"
        ):
            return

        if (
            self.listbox.identify_column(
                event.x
            )
            != "#5"
        ):
            return

        item = self.listbox.identify_row(
            event.y
        )

        if not item:
            return

        index = self.listbox.index(
            item
        )

        if index >= len(
            self._current_items
        ):
            return

        link = self._current_items[
            index
        ].get("link", "")

        if not link:
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(link)

        self.set_status("已复制")

        self._toast(
            event.x_root,
            event.y_root,
            "已复制",
        )

        return "break"

    # ========================================================
    # Detail
    # ========================================================

    def _clear_detail(self):

        self._detail_record = None

        self.detail_content.pack_forget()

        self.detail_empty.pack(
            fill=tk.BOTH,
            expand=True,
        )

    def show_record_detail(
        self,
        record,
    ):

        self._detail_record = record

        self.detail_empty.pack_forget()

        self.detail_content.pack(
            fill=tk.BOTH,
            expand=True,
        )

        self.detail_canvas.yview_moveto(0)

        for key, label in self.detail_labels.items():

            value = record.get(
                key,
                "",
            )

            label.config(
                text=value or "-",
                bg=C["card"],
                fg=C["sub"],
                padx=0,
                pady=0,
            )

        status = record.get(
            "status",
            "",
        )

        self.detail_labels[
            "status"
        ].config(
            bg=status_bg(status),
            fg=status_fg(status),
            padx=6,
            pady=3,
        )

    def copy_current_detail_link(self):

        if not self._detail_record:
            return

        link = self._detail_record.get(
            "link",
            "",
        )

        if not link:
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(link)

        self.set_status("已复制")

    # ========================================================
    # Helpers
    # ========================================================

    def _status_tag(self, status):

        return (
            "status_"
            + "".join(
                char
                if char.isalnum()
                else "_"
                for char in (
                    status
                    or "default"
                )
            )
        )

    def _toast(
        self,
        x,
        y,
        text,
    ):

        toast = tk.Toplevel(
            self.root
        )

        toast.overrideredirect(True)

        toast.attributes(
            "-topmost",
            True,
        )

        tk.Label(
            toast,
            text=text,

            bg=C["text"],
            fg=C["white"],

            font=(FONT, 9),

            padx=14,
            pady=7,
        ).pack()

        toast.update_idletasks()

        toast.geometry(
            f"+{x - toast.winfo_width() // 2}"
            f"+{y - 40}"
        )

        toast.after(
            800,
            toast.destroy,
        )

    def save_all(self):

        with JSONL_PATH.open(
            "w",
            encoding="utf-8",
        ) as file:

            for record in self.records:

                file.write(
                    json.dumps(
                        record,
                        ensure_ascii=False,
                    )
                    + "\n"
                )

    def _update_sidebar_count(self):

        if hasattr(
            self,
            "sidebar_count",
        ):
            self.sidebar_count.config(
                text=f"{len(self.records)} 条链接"
            )

    def set_status(
        self,
        text,
    ):

        self.statusbar.config(
            text=text
        )


# ============================================================
# Main
# ============================================================

def main():

    root = tk.Tk()

    LinkBrowserApp(root)

    root.mainloop()


if __name__ == "__main__":
    main()
