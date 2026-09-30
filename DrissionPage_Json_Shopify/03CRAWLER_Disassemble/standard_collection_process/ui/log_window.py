from __future__ import annotations

from tkinter import scrolledtext
import tkinter as tk
from ..ui.theme import MONO_FONT, UI_COLORS, UI_FONT


class LogWindow:
    """Only the Tk main thread may append log text."""

    def __init__(self, root):
        self.root = root
        self._build_log_window()

    def append(self, text):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

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
