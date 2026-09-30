from __future__ import annotations

from tkinter import messagebox, ttk
import os
import tkinter as tk
from ..ui.theme import MONO_FONT, UI_COLORS, UI_FONT


class UiComponents:
    """Reusable Tk controls and file-list rendering only."""

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
                UI_COLORS["danger_border"] if kind == "danger" else UI_COLORS["border"]
            ),
            highlightcolor=(
                UI_COLORS["danger_border"] if kind == "danger" else UI_COLORS["border"]
            ),
            cursor="hand2",
        )
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
            text="第二阶段 · 待处理文件",
            bg=UI_COLORS["card"],
            fg=UI_COLORS["text"],
            font=(UI_FONT, 11, "bold"),
        ).pack(anchor=tk.W)
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
            ("跳过图片位置", 2, 130, tk.W),
            ("保留图片位置", 3, 130, tk.W),
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

    def _populate_file_rows(self):
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showerror("路径无效", "请先选择一个有效的分类文件夹。")
            return

        xlsx_files = sorted(
            f for f in os.listdir(folder) if f.lower().endswith(".xlsx") and not f.startswith("~")
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
        self.list_inner.grid_columnconfigure(2, minsize=130)
        self.list_inner.grid_columnconfigure(3, minsize=100)

        for index, filename in enumerate(xlsx_files):
            file_path = os.path.join(folder, filename)
            skip_var = tk.StringVar()
            keep_var = tk.StringVar()
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

            skip_shell = tk.Frame(
                row,
                bg=row_bg,
                width=130,
            )
            skip_shell.grid(row=0, column=2, sticky="w", padx=(0, 10))
            skip_entry = tk.Entry(
                skip_shell,
                textvariable=skip_var,
                width=11,
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
                text="如 1,3,-1",
                bg=row_bg,
                fg=UI_COLORS["text_muted"],
                font=(UI_FONT, 8),
            ).pack(side=tk.LEFT, padx=(6, 0))

            keep_entry = tk.Entry(row, textvariable=keep_var, width=11, font=(UI_FONT, 9))
            keep_entry.grid(row=0, column=3, sticky="w", padx=(0, 10), ipady=5)
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
            "partial": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
            "stopped": (UI_COLORS["warning_light"], UI_COLORS["warning"]),
        }
        bg, fg = palettes.get(state, palettes["waiting"])
        row["status_var"].set(text)
        row["status_label"].configure(bg=bg, fg=fg)

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
