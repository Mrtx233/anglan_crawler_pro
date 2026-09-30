"""Shopify 价格修正工具

遍历 CSV 的每一行，当 Sale price 与 Regular price 同时有值且相等时，
把 Regular price 改为原值的 1.2 倍（保留两位小数），其余字段原样保留。

支持两种用法：
    1. 直接双击运行 -> 图形界面
    2. 命令行       -> python equal_price_fix.py <文件或文件夹> [--suffix 后缀] [--no-overwrite]
"""

import csv
import os
import sys
import tempfile
import threading
from pathlib import Path

import tkinter as tk
from tkinter import BooleanVar, StringVar, filedialog, messagebox
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

DEFAULT_SUFFIX = "_价格修正"
MARKUP = 1.2

SALE_COLUMN = "Sale price"
REGULAR_COLUMN = "Regular price"


# ---------- 通用工具 ----------


def strip_quotes(path):
    """去除路径两端可能包裹的引号（用户粘贴路径时可能带有引号）。"""
    path = str(path).strip()
    if len(path) >= 2 and (
            (path[0] == '"' and path[-1] == '"') or (path[0] == "'" and path[-1] == "'")
    ):
        path = path[1:-1].strip()
    return path


def detect_encoding(path):
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            with open(path, encoding=enc, newline="") as f:
                f.read(2048)
            return enc
        except UnicodeDecodeError:
            continue
    return "utf-8-sig"


def collect_csv_files(target):
    """目标可以是单个 csv 文件，也可以是一个文件夹（递归收集其中的 csv）。"""
    path = Path(strip_quotes(target))
    if path.is_dir():
        return sorted(p for p in path.rglob("*.csv") if p.is_file())
    if path.is_file():
        return [path]
    return []


def parse_price(value):
    """把单元格里的价格文本转成 float，无法解析时返回 None。"""
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("$", "").replace("￥", "").replace("¥", "").strip()
    try:
        return float(text)
    except ValueError:
        return None


def format_price(value):
    return f"{value:.2f}"


# ---------- 核心处理 ----------


def fix_equal_prices(csv_path, overwrite=True, suffix=DEFAULT_SUFFIX, log=print):
    """按规则修正单个 CSV，返回统计信息 dict。"""
    csv_path = Path(csv_path)
    encoding = detect_encoding(csv_path)

    with open(csv_path, encoding=encoding, newline="") as f:
        reader = csv.reader(f)
        try:
            header = next(reader)
        except StopIteration:
            log(f"[跳过] 空文件：{csv_path.name}")
            return None
        rows = list(reader)

    if SALE_COLUMN not in header or REGULAR_COLUMN not in header:
        log(f"[跳过] 缺少 {SALE_COLUMN} / {REGULAR_COLUMN} 列：{csv_path.name}")
        return None

    sale_idx = header.index(SALE_COLUMN)
    regular_idx = header.index(REGULAR_COLUMN)

    changed = 0
    for row in rows:
        if len(row) <= max(sale_idx, regular_idx):
            continue
        sale = parse_price(row[sale_idx])
        regular = parse_price(row[regular_idx])
        if sale is None or regular is None:
            continue
        if abs(sale - regular) > 1e-9:
            continue
        row[regular_idx] = format_price(regular * MARKUP)
        changed += 1

    if changed == 0:
        log(f"[完成] {csv_path.name}：没有需要修正的行（共 {len(rows)} 行）")
        return {"file": csv_path, "total": len(rows), "changed": 0, "output": csv_path}

    if overwrite:
        output_path = csv_path
    else:
        output_path = csv_path.with_name(f"{csv_path.stem}{suffix}{csv_path.suffix}")

    write_csv(output_path, header, rows, overwrite_from=csv_path if overwrite else None)

    log(f"[完成] {csv_path.name}：共 {len(rows)} 行，修正 {changed} 行 -> {output_path.name}")
    return {"file": csv_path, "total": len(rows), "changed": changed, "output": output_path}


def write_csv(output_path, header, rows, overwrite_from=None):
    """写出 utf-8-sig（带 BOM）的 CSV；覆盖原文件时先写临时文件再替换，避免中途失败损坏数据。"""
    output_path = Path(output_path)
    if overwrite_from is not None:
        fd, tmp_name = tempfile.mkstemp(dir=str(output_path.parent), suffix=".tmp")
        os.close(fd)
        try:
            _dump(tmp_name, header, rows)
            os.replace(tmp_name, output_path)
        except PermissionError:
            raise PermissionError(f"文件被其他程序占用（如 Excel），请关闭后重试：{output_path}") from None
        finally:
            if os.path.exists(tmp_name):
                os.remove(tmp_name)
    else:
        try:
            _dump(output_path, header, rows)
        except PermissionError:
            raise PermissionError(f"文件被其他程序占用（如 Excel），请关闭后重试：{output_path}") from None


def _dump(path, header, rows):
    # newline="" + 默认 lineterminator="\r\n"，与 Shopify 导出的 CSV 保持一致
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(header)
        writer.writerows(rows)


# ---------- 图形界面 ----------


class EqualPriceFixGui:
    def __init__(self, root):
        self.root = root
        self.root.title("Shopify 价格修正工具")
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
        tk.Label(header_frame, text="✨ Shopify 价格修正工具",
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
            text=(f"当「{SALE_COLUMN}」与「{REGULAR_COLUMN}」都有值且相等时，\n"
                  f"把「{REGULAR_COLUMN}」改为原值的 {MARKUP} 倍并保留两位小数。\n"
                  "其余字段、行序原样保留，结果按 utf-8（带 BOM）CSV 保存。"),
        ).pack(anchor="w")

        pick = ttk.LabelFrame(left_panel, text="  📂 选择目标  ", padding=12)
        pick.pack(fill="both", expand=True, pady=(15, 0))

        btn_row = ttk.Frame(pick)
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text="选择 CSV 文件", command=self.pick_files).pack(side="left")
        ttk.Button(btn_row, text="选择文件夹", command=self.pick_folder).pack(side="left", padx=8)
        ttk.Button(btn_row, text="清空列表", command=self.clear_files).pack(side="left")

        list_frame = ttk.Frame(pick)
        list_frame.pack(fill="both", expand=True, pady=(10, 0))
        self.file_list = tk.Listbox(list_frame, height=8, font=("Consolas", 9), activestyle="none")
        scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        self.file_list.configure(yscrollcommand=scroll.set)
        self.file_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        option = ttk.LabelFrame(left_panel, text="  ⚙️ 输出选项  ", padding=12)
        option.pack(fill="x", pady=(15, 0))

        self.overwrite_var = BooleanVar(value=True)
        ttk.Checkbutton(option, text="覆盖原文件（不勾选则另存为新文件）",
                        variable=self.overwrite_var).pack(anchor="w")

        suffix_row = ttk.Frame(option)
        suffix_row.pack(fill="x", pady=(8, 0))
        ttk.Label(suffix_row, text="另存后缀：").pack(side="left")
        self.suffix_var = StringVar(value=DEFAULT_SUFFIX)
        ttk.Entry(suffix_row, textvariable=self.suffix_var, width=18).pack(side="left")

        self.run_btn = ttk.Button(left_panel, text="🚀 开始修正", command=self.start)
        self.run_btn.pack(fill="x", pady=(15, 0))

        log_frame = ttk.LabelFrame(right_panel, text="  📝 运行日志  ", padding=12)
        log_frame.pack(fill="both", expand=True, padx=(10, 0))
        self.log_box = ScrolledText(log_frame, state="disabled", wrap="word", font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True)

    # ---------- 界面事件 ----------

    def pick_files(self):
        paths = filedialog.askopenfilenames(title="选择 CSV 文件", filetypes=[("CSV 文件", "*.csv"), ("全部文件", "*.*")])
        for path in paths:
            for csv_path in collect_csv_files(path):
                if csv_path not in self.files:
                    self.files.append(csv_path)
        self.refresh_list()

    def pick_folder(self):
        folder = filedialog.askdirectory(title="选择包含 CSV 的文件夹")
        if not folder:
            return
        found = collect_csv_files(folder)
        if not found:
            messagebox.showwarning("提示", "该文件夹下没有找到 csv 文件。")
            return
        for csv_path in found:
            if csv_path not in self.files:
                self.files.append(csv_path)
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
            messagebox.showwarning("提示", "请先选择要处理的 CSV 文件或文件夹。")
            return
        self.running = True
        self.run_btn.configure(state="disabled")
        self.status.set("⏳ 正在处理...")
        threading.Thread(target=self.run, daemon=True).start()

    def run(self):
        overwrite = self.overwrite_var.get()
        suffix = self.suffix_var.get().strip() or DEFAULT_SUFFIX
        total_rows = total_changed = failed = 0
        try:
            self.log(f"开始处理 {len(self.files)} 个文件（{'覆盖原文件' if overwrite else '另存为新文件'}）")
            self.log("-" * 60)
            for path in list(self.files):
                try:
                    stat = fix_equal_prices(path, overwrite=overwrite, suffix=suffix, log=self.log)
                except Exception as exc:  # 单个文件失败不影响后续文件
                    failed += 1
                    self.log(f"[失败] {path.name}：{exc}")
                    continue
                if stat:
                    total_rows += stat["total"]
                    total_changed += stat["changed"]
            self.log("-" * 60)
            summary = f"处理完成：{len(self.files) - failed} 个文件，共 {total_rows} 行，修正 {total_changed} 行"
            if failed:
                summary += f"，失败 {failed} 个"
            self.log(summary)
            self.status.set("✅ " + summary)
        finally:
            self.running = False
            self.root.after(0, lambda: self.run_btn.configure(state="normal"))


# ---------- 命令行入口 ----------


def run_cli(argv):
    targets = [a for a in argv if not a.startswith("--")]
    if not targets:
        return False

    overwrite = "--no-overwrite" not in argv
    suffix = DEFAULT_SUFFIX
    if "--suffix" in argv:
        idx = argv.index("--suffix")
        if idx + 1 < len(argv):
            suffix = argv[idx + 1]

    files = []
    for target in targets:
        found = collect_csv_files(target)
        if not found:
            print(f"[跳过] 找不到目标：{target}")
        files.extend(found)

    total_rows = total_changed = failed = 0
    for path in files:
        try:
            stat = fix_equal_prices(path, overwrite=overwrite, suffix=suffix)
        except Exception as exc:  # 单个文件失败不影响后续文件
            failed += 1
            print(f"[失败] {path.name}：{exc}")
            continue
        if stat:
            total_rows += stat["total"]
            total_changed += stat["changed"]
    summary = f"共处理 {len(files) - failed} 个文件，{total_rows} 行，修正 {total_changed} 行"
    if failed:
        summary += f"，失败 {failed} 个"
    print(summary)
    return True


def main():
    argv = sys.argv[1:]
    if argv and run_cli(argv):
        return

    root = tk.Tk()
    EqualPriceFixGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
