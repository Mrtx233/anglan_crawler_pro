# -*- coding: utf-8 -*-
"""按 title 分组拆分 Excel 工具

读取 xlsx，按指定列（默认 title）分组，组内按价格列（默认 price1）升序排序，
再按轮询方式（第 i 行 -> 第 i 个文件）均匀分配到 N 个文件中。
同一款式的不同价格会分散到各个文件，避免全部挤在一起。

输出到源文件所在目录，命名为「源文件名_1.xlsx」「源文件名_2.xlsx」…

直接双击运行即可打开图形界面。
"""

import os
import sys
import threading
from pathlib import Path

import openpyxl
import tkinter as tk
from tkinter import BooleanVar, IntVar, StringVar, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

DEFAULT_TITLE_COL = "title"
DEFAULT_PRICE_COL = "price1"
DEFAULT_PARTS = 2


# ---------- 通用工具 ----------


def strip_quotes(path):
    """去除路径两端可能包裹的引号（用户粘贴路径时可能带有引号）。"""
    path = str(path).strip()
    if len(path) >= 2 and (
            (path[0] == '"' and path[-1] == '"') or (path[0] == "'" and path[-1] == "'")
    ):
        path = path[1:-1].strip()
    return path


def collect_xlsx_files(target):
    """目标可以是单个 xlsx 文件，也可以是一个文件夹（递归收集其中的 xlsx）。"""
    path = Path(strip_quotes(target))
    if path.is_dir():
        return sorted(
            p for p in path.rglob("*.xlsx")
            if p.is_file() and not p.name.startswith("~$")
        )
    if path.is_file() and path.suffix.lower() == ".xlsx":
        return [path]
    return []


def parse_price(value):
    """把单元格里的价格转成 float，无法解析时返回 None。"""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("$", "").replace("￥", "").replace("¥", "").strip()
    try:
        return float(text)
    except ValueError:
        return None


def price_sort_key(value):
    """价格升序，无法解析的值排到末尾。"""
    price = parse_price(value)
    return (price is None, price if price is not None else 0.0)


# ---------- 核心处理 ----------


def split_by_title(xlsx_path, parts, title_col=DEFAULT_TITLE_COL, price_col=DEFAULT_PRICE_COL,
                   overwrite=True, log=print):
    """按 title_col 分组、按 price_col 排序后轮询分配到 parts 个 xlsx，返回统计信息 dict。"""
    xlsx_path = Path(xlsx_path)
    parts = max(1, int(parts))

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    try:
        header = list(next(rows_iter))
    except StopIteration:
        wb.close()
        log(f"[跳过] 空文件：{xlsx_path.name}")
        return None
    rows = [list(r) for r in rows_iter]
    wb.close()

    if title_col not in header:
        log(f"[跳过] 缺少「{title_col}」列：{xlsx_path.name}")
        return None

    ti = header.index(title_col)
    has_price = price_col in header
    if not has_price:
        log(f"[提示] 缺少「{price_col}」列，组内保持原顺序：{xlsx_path.name}")
    pi = header.index(price_col) if has_price else None

    groups = {}
    for r in rows:
        groups.setdefault(r[ti], []).append(r)
    if pi is not None:
        for g in groups.values():
            g.sort(key=lambda r: price_sort_key(r[pi]))

    buckets = [[] for _ in range(parts)]
    for g in groups.values():
        for i, r in enumerate(g):
            buckets[i % parts].append(r)

    stem = xlsx_path.stem
    outdir = xlsx_path.parent
    outputs = []
    for k, bucket in enumerate(buckets, 1):
        out_path = outdir / f"{stem}_{k}.xlsx"
        if not overwrite and out_path.exists():
            out_path = outdir / f"{stem}_{k}_新.xlsx"
        save_workbook(out_path, header, bucket)
        outputs.append(out_path)
        log(f"[完成] {out_path.name} -> {len(bucket)} 行")

    log(f"[完成] {xlsx_path.name}：共 {len(rows)} 行 / {len(groups)} 组，拆分 {parts} 份")
    return {
        "file": xlsx_path,
        "total": len(rows),
        "groups": len(groups),
        "parts": parts,
        "outputs": outputs,
    }


def save_workbook(path, header, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(header)
    for r in rows:
        ws.append(r)
    try:
        wb.save(path)
    except PermissionError:
        raise PermissionError(f"文件被其他程序占用（如 Excel），请关闭后重试：{path}") from None


# ---------- 图形界面 ----------


class SplitByTitleGui:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel 按 title 拆分工具")
        self.files = []
        self.running = False

        window_width = 960
        window_height = 820
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(860, 540)

        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        if sys.platform.startswith("win"):
            style.theme_use("vista")
        style.configure(".", font=("Microsoft YaHei UI", 10))
        style.configure("TButton", padding=6)
        style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 11, "bold"), foreground="#0078D4")

        bottom = tk.Frame(self.root, bg="#E8F4F8")
        bottom.pack(side="bottom", fill="x")
        self.status = StringVar(value="✨ 就绪")
        tk.Label(bottom, textvariable=self.status, font=("Microsoft YaHei UI", 10),
                 bg="#E8F4F8", fg="#333333").pack(side="left", padx=20, pady=8)

        outer = ttk.Frame(self.root)
        outer.pack(fill="both", expand=True)

        header_frame = tk.Frame(outer, bg="#0078D4")
        header_frame.pack(fill="x")
        tk.Label(header_frame, text="✨ Excel 按 title 拆分工具",
                 font=("Microsoft YaHei UI", 16, "bold"), bg="#0078D4", fg="#FFFFFF").pack(side="left", padx=20, pady=15)

        paned = ttk.Panedwindow(outer, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=20, pady=20)

        left_panel = ttk.Frame(paned)
        paned.add(left_panel, weight=11)

        right_panel = ttk.Frame(paned)
        paned.add(right_panel, weight=9)

        tip = ttk.LabelFrame(left_panel, text="  📌 处理规则  ", padding=12)
        tip.pack(fill="x")
        ttk.Label(
            tip,
            justify="left",
            text=("按「分组列」把行归组，组内按「价格列」升序排序后，\n"
                  "轮询分配（第 1 行 -> 第 1 个文件，第 2 行 -> 第 2 个文件…），\n"
                  "使同一款式的不同价格尽量分散到各文件。\n"
                  "结果保存到源文件同目录，命名为「源文件名_1.xlsx」「源文件名_2.xlsx」…"),
        ).pack(anchor="w")

        pick = ttk.LabelFrame(left_panel, text="  📂 选择目标  ", padding=12)
        pick.pack(fill="both", expand=True, pady=(15, 0))

        btn_row = ttk.Frame(pick)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="选择 Excel 文件", command=self.pick_files).pack(side="left")
        ttk.Button(btn_row, text="选择文件夹", command=self.pick_folder).pack(side="left", padx=8)
        ttk.Button(btn_row, text="清空列表", command=self.clear_files).pack(side="left")

        list_frame = ttk.Frame(pick)
        list_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.file_list = tk.Listbox(list_frame, height=8, font=("Consolas", 9), activestyle="none")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scroll.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        option = ttk.LabelFrame(left_panel, text="  ⚙️ 拆分选项  ", padding=12)
        option.pack(fill="x", pady=(15, 0))

        col_row = ttk.Frame(option)
        col_row.pack(fill="x")
        ttk.Label(col_row, text="分组列：").pack(side="left")
        self.title_col_var = StringVar(value=DEFAULT_TITLE_COL)
        ttk.Entry(col_row, textvariable=self.title_col_var, width=15).pack(side="left")
        ttk.Label(col_row, text="    价格列：").pack(side="left")
        self.price_col_var = StringVar(value=DEFAULT_PRICE_COL)
        ttk.Entry(col_row, textvariable=self.price_col_var, width=15).pack(side="left")

        parts_row = ttk.Frame(option)
        parts_row.pack(fill="x", pady=(8, 0))
        ttk.Label(parts_row, text="拆分份数：").pack(side="left")
        self.parts_var = IntVar(value=DEFAULT_PARTS)
        ttk.Spinbox(parts_row, from_=1, to=99, width=5, textvariable=self.parts_var).pack(side="left")
        ttk.Label(parts_row, text="    （份数大于组内行数时，多出的文件为空）",
                  foreground="#888888").pack(side="left")

        self.overwrite_var = BooleanVar(value=True)
        ttk.Checkbutton(option, text="覆盖同名文件（不勾选则加「_新」另存）",
                        variable=self.overwrite_var).pack(anchor="w", pady=(8, 0))

        self.run_btn = ttk.Button(left_panel, text="🚀 开始拆分", command=self.start)
        self.run_btn.pack(fill="x", pady=(15, 0))

        log_frame = ttk.LabelFrame(right_panel, text="  📝 运行日志  ", padding=12)
        log_frame.pack(fill="both", expand=True, padx=(10, 0))
        self.log_box = ScrolledText(log_frame, state="disabled", wrap="word", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True)

    # ---------- 界面事件 ----------

    def pick_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("全部文件", "*.*")],
        )
        for path in paths:
            for xlsx_path in collect_xlsx_files(path):
                if xlsx_path not in self.files:
                    self.files.append(xlsx_path)
        self.refresh_list()

    def pick_folder(self):
        folder = filedialog.askdirectory(title="选择包含 Excel 的文件夹")
        if not folder:
            return
        found = collect_xlsx_files(folder)
        if not found:
            messagebox.showwarning("提示", "该文件夹下没有找到 xlsx 文件。")
            return
        for xlsx_path in found:
            if xlsx_path not in self.files:
                self.files.append(xlsx_path)
        self.refresh_list()

    def clear_files(self):
        self.files = []
        self.refresh_list()

    def refresh_list(self):
        self.file_list.delete(0, tk.END)
        for path in self.files:
            self.file_list.insert(tk.END, str(path))
        self.status.set(f"✨ 已选择 {len(self.files)} 个文件")

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert(tk.END, message + "\n")
        self.log_box.see(tk.END)
        self.log_box.configure(state="disabled")

    def start(self):
        if self.running:
            return
        if not self.files:
            messagebox.showwarning("提示", "请先选择要拆分的 Excel 文件或文件夹。")
            return
        try:
            parts = int(self.parts_var.get())
        except (tk.TclError, ValueError):
            messagebox.showwarning("提示", "拆分份数必须是数字。")
            return
        if parts < 1:
            messagebox.showwarning("提示", "拆分份数至少为 1。")
            return
        if not self.title_col_var.get().strip():
            messagebox.showwarning("提示", "请填写分组列名。")
            return

        self.running = True
        self.run_btn.configure(state="disabled")
        self.status.set("⏳ 正在处理...")
        threading.Thread(target=self.run, args=(parts,), daemon=True).start()

    def run(self, parts):
        title_col = self.title_col_var.get().strip()
        price_col = self.price_col_var.get().strip()
        overwrite = self.overwrite_var.get()
        total_rows = total_files = failed = 0
        try:
            self.log(f"开始处理 {len(self.files)} 个文件，拆分 {parts} 份（{'覆盖同名文件' if overwrite else '加「_新」另存'}）")
            self.log(f"分组列：{title_col}    价格列：{price_col or '（空，保持原序）'}")
            self.log("-" * 60)
            for path in list(self.files):
                try:
                    stat = split_by_title(path, parts, title_col, price_col,
                                          overwrite=overwrite, log=self.log)
                except Exception as exc:  # 单个文件失败不影响后续文件
                    failed += 1
                    self.log(f"[失败] {path.name}：{exc}")
                    continue
                if stat:
                    total_files += 1
                    total_rows += stat["total"]
            self.log("-" * 60)
            summary = f"处理完成：{total_files} 个文件，共 {total_rows} 行，各拆成 {parts} 份"
            if failed:
                summary += f"，失败 {failed} 个"
            self.log(summary)
            self.status.set("✅ " + summary)
        finally:
            self.running = False
            self.root.after(0, lambda: self.run_btn.configure(state="normal"))


def main():
    root = tk.Tk()
    SplitByTitleGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
#（注：内容由AI生成）
