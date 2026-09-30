# ======================== GUI 入口 ========================
#
# 职责: 只负责三件事——画界面、校验输入组装 CollectConfig、
#       起线程并把进度/日志/保存报告转发回界面。不含采集与文件逻辑。
# 依赖: config、logger、collector（唯一允许导入 tkinter 的模块）。
#
# 布局: 卡片式单列——「1 保存位置」→「2 分类列表」→「可选设置」→ 页脚
#       （状态 / 链接计数 / 按钮 / 进度条 / 说明）。日志独立成窗，默认隐藏，
#       由页脚的「查看日志」唤起。
#
# 运行: cd link_collector_pagination && python main.py
# ========================================================

import os
import threading
import time

import tkinter as tk
from tkinter import filedialog, font as tkfont, messagebox, scrolledtext, ttk

# 按文档运行方式（在本目录执行 python main.py）时脚本目录已在 sys.path 上，
# 此处不再修改 sys.path，避免把自己的 config/logger 抢到同名模块之前。
import collector
import config
import logger

PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
)

# 关窗等待保存的最长秒数；超时则强制退出并打日志
CLOSE_WAIT_TIMEOUT = 60
# 轮询采集线程的间隔（毫秒）
CLOSE_POLL_INTERVAL_MS = 200


def pick_font(candidates, fallback):
    """取第一个系统已安装的候选字体；都没有时用 fallback。"""
    families = set(tkfont.families())
    for name in candidates:
        if name in families:
            return name
    return fallback


class LinkCollectorApp:
    def __init__(self, root):
        self.root = root
        self.is_running = False
        self.stop_event = threading.Event()
        self.runner = None
        self._worker_thread = None
        self._closing = False
        # 采集进行中需要一并禁用的输入控件
        self._editable = []
        # 进度条总量，收尾时用来把进度补满
        self._progress_max = 1

        self.max_pages_var = tk.StringVar(value="")
        self.xpath_var = tk.StringVar()
        self.output_var = tk.StringVar(
            value=os.path.join(PROJECT_ROOT, config.DEFAULT_OUTPUT_DIRNAME)
        )
        self.status_var = tk.StringVar(value=config.STATUS_READY)
        self.count_var = tk.StringVar(value="0 条链接")
        self.category_count_var = tk.StringVar(value="0 个分类")
        self.summary_var = tk.StringVar(value=config.FOOTER_HINT)

        self.root.title(config.APP_TITLE)
        self.root.geometry(config.WINDOW_GEOMETRY)
        self.root.minsize(*config.WINDOW_MIN_SIZE)
        self.root.configure(bg=config.UI_COLORS["background"])

        self._build_ui()

        self._log_sink = self._append_log
        logger.add_sink(self._log_sink)
        logger.install_stdout_redirect()

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ======================== 界面构建 ========================
    def _build_ui(self):
        self.family = pick_font(config.UI_FONT_CANDIDATES, config.UI_FONT_FALLBACK)
        self.mono = pick_font(config.MONO_FONT_CANDIDATES, config.MONO_FONT_FALLBACK)

        self._install_styles()
        self._build_log_window()

        body = ttk.Frame(self.root, padding=16)
        body.pack(fill=tk.BOTH, expand=True)
        body.columnconfigure(0, weight=1)
        # 分类列表所在行吸收多余高度，其余行保持自然高度
        body.rowconfigure(2, weight=1)

        self._build_head(body)
        self._build_output_row(body)
        self._build_categories(body)
        self._build_options(body)
        self._build_footer(body)

        # 初始焦点放在分类输入框：既是用户第一个要填的地方，
        # 也避免按钮抢先获得焦点时带上虚线框。
        self.pages_text.focus_set()

    def _install_styles(self):
        """集中注册 ttk 样式；配色全部取自 config.UI_COLORS。"""
        colors = config.UI_COLORS
        family = self.family

        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=colors["background"])
        style.configure("Card.TFrame", background=colors["card"])
        style.configure(
            "TLabel",
            background=colors["card"],
            foreground=colors["text"],
            font=(family, 11),
        )
        style.configure(
            "Muted.TLabel", foreground=colors["muted"], font=(family, 10)
        )
        style.configure("Section.TLabel", font=(family, 12, "bold"))
        style.configure(
            "TEntry",
            padding=7,
            fieldbackground=colors["card_alt"],
            font=(family, 11),
        )
        style.configure("TButton", padding=(14, 9), font=(family, 11))
        style.configure(
            "Primary.TButton",
            background=colors["primary"],
            foreground=colors["card"],
        )
        style.map(
            "Primary.TButton",
            background=[
                ("disabled", colors["primary_disabled"]),
                ("active", colors["primary_hover"]),
            ],
        )
        style.configure("Danger.TButton", foreground=colors["danger"])
        style.configure(
            "TProgressbar",
            background=colors["primary"],
            troughcolor=colors["progress_trough"],
            thickness=5,
        )

    def _build_head(self, body):
        head = ttk.Frame(body)
        head.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        tk.Label(
            head, text=config.APP_HEADING,
            font=(self.family, 18, "bold"),
            bg=config.UI_COLORS["background"], fg=config.UI_COLORS["text"],
        ).pack(anchor=tk.W)

    def _build_output_row(self, body):
        row = ttk.Frame(body, style="Card.TFrame", padding=12)
        row.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        row.columnconfigure(1, weight=1)

        ttk.Label(row, text="1  保存位置", style="Section.TLabel").grid(
            row=0, column=0, padx=(0, 16)
        )
        path_entry = ttk.Entry(row, textvariable=self.output_var)
        path_entry.grid(row=0, column=1, sticky="ew")
        browse_btn = ttk.Button(
            row, text="选择文件夹", command=self._browse_output
        )
        browse_btn.grid(row=0, column=2, padx=(10, 0))

        self._editable.extend([path_entry, browse_btn])

    def _build_categories(self, body):
        colors = config.UI_COLORS
        card = ttk.Frame(body, style="Card.TFrame", padding=14)
        card.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        card.columnconfigure(0, weight=1)
        card.rowconfigure(2, weight=1)

        ttk.Label(card, text="2  分类列表", style="Section.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            card, textvariable=self.category_count_var, style="Muted.TLabel"
        ).grid(row=0, column=1, sticky="e")
        ttk.Label(
            card, text=config.CATEGORY_PLACEHOLDER, style="Muted.TLabel"
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 10))

        self.pages_text = scrolledtext.ScrolledText(
            card, height=18, wrap=tk.NONE,
            bg=colors["card_alt"], fg=colors["text"],
            insertbackground=colors["primary"],
            font=(self.mono, 11), relief=tk.FLAT, bd=0,
            highlightthickness=1, highlightbackground=colors["border"],
            highlightcolor=colors["primary"],
            padx=10, pady=8,
        )
        self.pages_text.grid(row=2, column=0, columnspan=2, sticky="nsew")
        self.pages_text.bind("<<Modified>>", self._update_category_count)

        self._editable.append(self.pages_text)

    def _build_options(self, body):
        row = ttk.Frame(body, style="Card.TFrame", padding=(14, 6))
        row.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        row.columnconfigure(4, weight=1)

        ttk.Label(row, text="可选设置", style="Muted.TLabel").grid(
            row=0, column=0, padx=(0, 16)
        )
        ttk.Label(row, text="每分类页数", style="Muted.TLabel").grid(
            row=0, column=1, padx=(0, 6)
        )
        limit_entry = ttk.Entry(row, textvariable=self.max_pages_var, width=5)
        limit_entry.grid(row=0, column=2)
        ttk.Label(row, text="XPath", style="Muted.TLabel").grid(
            row=0, column=3, padx=(16, 6)
        )
        xpath_entry = ttk.Entry(row, textvariable=self.xpath_var, width=24)
        xpath_entry.grid(row=0, column=4, sticky="ew")
        ttk.Label(row, text=config.OPTIONS_HINT, style="Muted.TLabel").grid(
            row=0, column=5, padx=(12, 0)
        )

        self._editable.extend([limit_entry, xpath_entry])

    def _build_footer(self, body):
        row = ttk.Frame(body, style="Card.TFrame", padding=10)
        row.grid(row=4, column=0, sticky="ew")
        row.columnconfigure(0, weight=1)

        ttk.Label(row, textvariable=self.status_var, wraplength=280).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            row, textvariable=self.count_var, style="Section.TLabel"
        ).grid(row=0, column=1, padx=16)

        ttk.Button(row, text="查看日志", command=self._show_log).grid(
            row=0, column=2, padx=(0, 8)
        )
        self.stop_btn = ttk.Button(
            row, text="停止并保存", command=self._stop,
            state=tk.DISABLED, style="Danger.TButton",
        )
        self.stop_btn.grid(row=0, column=3, padx=(0, 8))
        self.start_btn = ttk.Button(
            row, text="开始采集", command=self._start, style="Primary.TButton"
        )
        self.start_btn.grid(row=0, column=4)

        self.progress = ttk.Progressbar(row, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=5, sticky="ew", pady=(10, 7))
        ttk.Label(
            row, textvariable=self.summary_var, style="Muted.TLabel",
            wraplength=830,
        ).grid(row=2, column=0, columnspan=5, sticky="w")

    def _build_log_window(self):
        """日志独立成窗；关闭只是隐藏，采集线程仍持续写入并保留历史。"""
        colors = config.UI_COLORS
        self.log_window = tk.Toplevel(self.root)
        self.log_window.withdraw()
        self.log_window.title(config.LOG_WINDOW_TITLE)
        self.log_window.geometry(config.LOG_WINDOW_GEOMETRY)
        self.log_window.minsize(*config.LOG_WINDOW_MIN_SIZE)
        self.log_window.configure(bg=colors["background"])
        self.log_window.protocol("WM_DELETE_WINDOW", self.log_window.withdraw)

        panel = ttk.Frame(self.log_window, style="Card.TFrame", padding=14)
        panel.pack(fill=tk.BOTH, expand=True)
        panel.columnconfigure(0, weight=1)
        panel.rowconfigure(1, weight=1)

        ttk.Label(panel, text="运行日志", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 10)
        )
        ttk.Button(panel, text="清空日志", command=self._clear_log).grid(
            row=0, column=1, pady=(0, 10)
        )
        self.log_text = scrolledtext.ScrolledText(
            panel, state=tk.DISABLED,
            bg=colors["log_background"], fg=colors["log_text"],
            insertbackground=colors["log_text"],
            font=(self.mono, 10), relief=tk.FLAT, bd=0, wrap=tk.WORD,
            padx=12, pady=10,
        )
        self.log_text.grid(row=1, column=0, columnspan=2, sticky="nsew")

    # ======================== 界面回调 ========================
    def _show_log(self):
        self.log_window.deiconify()
        self.log_window.lift()
        self.log_text.focus_set()

    def _browse_output(self):
        path = filedialog.askdirectory(
            title="选择保存目录", initialdir=self.output_var.get() or None
        )
        if path:
            self.output_var.set(path)

    def _update_category_count(self, event=None):
        if self.pages_text.edit_modified():
            count = len(
                collector.parse_categories(self.pages_text.get("1.0", tk.END))
            )
            self.category_count_var.set(f"{count} 个分类")
            self.pages_text.edit_modified(False)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _append_log(self, text):
        if not text:
            return
        try:
            self.root.after(0, self._write_log, text)
        except (tk.TclError, RuntimeError):
            pass

    def _write_log(self, text):
        try:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        except tk.TclError:
            pass

    def _on_progress(self, index, total, title):
        """分类进度回调（采集线程）→ 主线程更新进度条与状态。"""

        def _apply():
            self._progress_max = max(1, total)
            self.progress.configure(maximum=self._progress_max, value=index - 1)
            if not self.stop_event.is_set():
                self.status_var.set(f"分类 {index}/{total} · {title[:24]}")

        try:
            self.root.after(0, _apply)
        except (tk.TclError, RuntimeError):
            pass

    def _on_count(self, count):
        """每提交一页后回报累计链接数（采集线程）。"""
        try:
            self.root.after(0, self.count_var.set, f"{count} 条链接")
        except (tk.TclError, RuntimeError):
            pass

    def _on_save_report(self, report):
        """采集线程的保存报告回调，转发到主线程更新摘要/弹窗。"""

        def _apply():
            text = report.describe()
            self.summary_var.set(text)
            self.count_var.set(f"{report.total_rows} 条链接")
            self.progress.configure(value=self._progress_max)
            self.status_var.set(self._final_status(report))

            # 关窗过程中不弹窗，只写日志；避免阻塞退出
            if self._closing:
                logger.log(f"[保存摘要] {text}")
                return

            if report.error:
                messagebox.showerror("保存失败", text)
                return

            notes = []
            if report.recovery_files:
                notes.append("恢复文件:\n" + "\n".join(report.recovery_files))
            if report.json_recovery_files:
                notes.append("索引备份:\n" + "\n".join(report.json_recovery_files))
            if report.index_error:
                notes.append(
                    f"索引更新失败（域名文件不受影响）:\n{report.index_error}"
                )
            if notes:
                messagebox.showwarning(
                    "保存完成（含异常提示）", "\n\n".join([text] + notes)
                )
            else:
                logger.log(f"[保存摘要] {text}")

        try:
            self.root.after(0, _apply)
        except (tk.TclError, RuntimeError):
            pass

    def _final_status(self, report):
        """把保存报告归纳成页脚状态文字。"""
        if report.error:
            return "保存失败 · 请查看运行日志"
        if self.stop_event.is_set():
            return "已停止并保存"
        if report.recovery_files or report.index_error:
            return "采集完成 · 保存存在异常提示"
        if not report.total_rows:
            return "采集结束 · 未找到商品链接"
        return "采集完成"

    # ======================== 输入校验与组装 ========================
    def _build_config(self):
        raw_text = self.pages_text.get("1.0", tk.END)
        categories = collector.parse_categories(raw_text)
        if not categories:
            messagebox.showerror("缺少分类", "请输入至少一个有效的分类 URL。")
            return None

        output_path = self.output_var.get().strip()
        if not output_path:
            output_path = os.path.join(
                PROJECT_ROOT, config.DEFAULT_OUTPUT_DIRNAME
            )
            self.output_var.set(output_path)

        max_pages_raw = self.max_pages_var.get().strip()
        if max_pages_raw:
            try:
                max_pages = int(max_pages_raw)
                if max_pages <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror(
                    "页数无效", "最大页数应为正整数，留空表示不限。"
                )
                return None
        else:
            max_pages = 0

        return config.CollectConfig(
            categories=categories,
            output_dir=output_path,
            max_pages=max_pages,
            xpath=self.xpath_var.get().strip(),
        )

    # ======================== 任务控制 ========================
    def _set_running(self, running):
        """切换运行态：禁用/恢复输入控件与开始按钮，反转停止按钮。

        只改控件状态，不动状态栏文字——收尾状态由保存报告回调负责，
        两者都在主线程排队执行，顺序上报告在前。
        """
        state = tk.DISABLED if running else tk.NORMAL
        for widget in self._editable:
            try:
                widget.configure(state=state)
            except tk.TclError:
                pass
        self.start_btn.configure(state=state)
        self.stop_btn.configure(state=tk.NORMAL if running else tk.DISABLED)

    def _start(self):
        if self.is_running:
            return

        cfg = self._build_config()
        if cfg is None:
            return

        self.is_running = True
        self._closing = False
        self.stop_event.clear()
        self._set_running(True)
        self._clear_log()
        self.status_var.set(config.STATUS_STARTING)
        self.count_var.set("0 条链接")
        self.summary_var.set(config.FOOTER_HINT)
        self._progress_max = max(1, len(cfg.categories))
        self.progress.configure(maximum=self._progress_max, value=0)

        self.runner = collector.CollectorRunner(
            cfg,
            stop_event=self.stop_event,
            on_progress=self._on_progress,
            on_save_report=self._on_save_report,
            on_count=self._on_count,
        )
        self._worker_thread = threading.Thread(
            target=self._worker, daemon=True, name="collector-worker"
        )
        self._worker_thread.start()

    def _worker(self):
        try:
            self.runner.run()
        finally:
            self.is_running = False
            if not self._closing:
                try:
                    self.root.after(0, self._set_running_done)
                except (tk.TclError, RuntimeError):
                    pass

    def _set_running_done(self):
        self._set_running(False)

    def _stop(self):
        if not self.is_running or self.runner is None:
            return
        self.stop_btn.config(state=tk.DISABLED)
        self.status_var.set(config.STATUS_STOPPING)
        logger.log("\n[停止] 已放弃当前页，正在保存此前已采集的链接...")
        threading.Thread(target=self.runner.interrupt, daemon=True).start()

    # ======================== 关闭与退出 ========================
    def _on_close(self):
        """关闭窗口：运行中先确认，再请求停止并等待保存完成后销毁。"""
        if self._closing:
            return

        if self.is_running:
            if not messagebox.askyesno(
                "确认退出",
                "采集任务正在进行中。\n"
                "退出将停止采集并保存此前已完成分页的结果，"
                "当前正在加载的页面会被放弃。\n\n确定退出吗？",
            ):
                return

            self._closing = True
            self.start_btn.config(state=tk.DISABLED)
            self.stop_btn.config(state=tk.DISABLED)
            self.status_var.set(config.STATUS_STOPPING)
            logger.log("\n[关闭] 正在停止采集并等待保存完成...")

            try:
                self.runner.interrupt()
            except Exception:
                pass

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
            logger.log(
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
        """还原 stdout 并注销日志 sink，避免窗口销毁后写入失效控件。"""
        try:
            logger.remove_sink(self._log_sink)
        finally:
            logger.restore_stdout()
        try:
            self.root.destroy()
        except tk.TclError:
            pass


def main():
    root = tk.Tk()
    LinkCollectorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
