import os
import sys
import glob
import csv
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from collections import Counter

# ---------- 自动安装依赖 ----------
_deps = {"Pillow": "PIL"}
for pkg, mod in _deps.items():
    try:
        __import__(mod)
    except ImportError:
        print(f"正在安装 {pkg}...")
        os.system(f'"{sys.executable}" -m pip install {pkg} -q')

from PIL import Image


class TextRedirector:
    """用于将 print() 输出重定向到 GUI 文本框的工具类"""

    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        # 使用 after 确保线程安全的 GUI 更新
        self.widget.after(0, self._write, text)

    def _write(self, text):
        self.widget.config(state=tk.NORMAL)
        self.widget.insert(tk.END, text)
        self.widget.see(tk.END)
        self.widget.config(state=tk.DISABLED)

    def flush(self):
        pass


class ImageConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("图片批量转 JPG 工具")
        self.geometry("650x550")
        self.resizable(True, True)

        # 核心格式配置
        self.HEIF_EXTS = {".heic", ".heif"}
        self.CONVERT_EXTS = {".webp", ".bmp", ".tiff", ".tif", ".gif", ".avif"}

        self.setup_ui()

        # 将标准输出重定向到日志框
        sys.stdout = TextRedirector(self.log_text)

    def setup_ui(self):
        # 配置全局间距
        padding = {'padx': 10, 'pady': 5}

        # --- 顶部路径配置区 ---
        frame_top = tk.Frame(self)
        frame_top.pack(fill=tk.X, padx=10, pady=10)

        # 图片文件夹路径
        tk.Label(frame_top, text="图片文件夹:").grid(row=0, column=0, sticky="e", **padding)
        self.entry_images_dir = ttk.Entry(frame_top, width=50)
        self.entry_images_dir.grid(row=0, column=1, sticky="we", **padding)
        ttk.Button(frame_top, text="浏览...", command=self.browse_images_dir).grid(row=0, column=2, **padding)

        # CSV 文件路径
        tk.Label(frame_top, text="CSV 文件(可选):").grid(row=1, column=0, sticky="e", **padding)
        self.entry_csv_file = ttk.Entry(frame_top, width=50)
        self.entry_csv_file.grid(row=1, column=1, sticky="we", **padding)
        ttk.Button(frame_top, text="浏览...", command=self.browse_csv_file).grid(row=1, column=2, **padding)

        # 设置列宽自适应
        frame_top.columnconfigure(1, weight=1)

        # --- 中部按钮区 ---
        frame_mid = tk.Frame(self)
        frame_mid.pack(fill=tk.X, padx=10, pady=5)

        self.btn_start = ttk.Button(frame_mid, text="▶ 开始转换", command=self.start_conversion)
        self.btn_start.pack(side=tk.LEFT, padx=5)

        self.btn_clear = ttk.Button(frame_mid, text="清空日志", command=self.clear_log)
        self.btn_clear.pack(side=tk.LEFT, padx=5)

        # --- 日志显示区 ---
        frame_log = tk.LabelFrame(self, text="处理日志")
        frame_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.log_text = scrolledtext.ScrolledText(frame_log, state=tk.DISABLED, wrap=tk.WORD, font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # --- 底部进度条 ---
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, padx=10, pady=10)

    def browse_images_dir(self):
        folder = filedialog.askdirectory(title="选择图片文件夹")
        if folder:
            self.entry_images_dir.delete(0, tk.END)
            self.entry_images_dir.insert(0, folder)

    def browse_csv_file(self):
        file = filedialog.askopenfilename(
            title="选择 CSV 文件",
            filetypes=(("CSV 文件", "*.csv"), ("所有文件", "*.*"))
        )
        if file:
            self.entry_csv_file.delete(0, tk.END)
            self.entry_csv_file.insert(0, file)

    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _register_heif(self):
        try:
            from pillow_heif import register_heif_opener
            register_heif_opener()
            return True
        except ImportError:
            print("[警告] 未安装 pillow-heif，.heic/.heif 无法转换（其他格式不受影响）。")
        except Exception as e:
            print(f"[警告] 无法加载 pillow-heif（.heic/.heif 将跳过）：{e}")
        return False

    def start_conversion(self):
        images_dir = self.entry_images_dir.get().strip()
        csv_file = self.entry_csv_file.get().strip()

        if not images_dir or not os.path.isdir(images_dir):
            print("[错误] 请选择有效的图片文件夹！")
            return

        # 禁用按钮，防止重复点击
        self.btn_start.config(state=tk.DISABLED)
        self.progress_var.set(0)
        print("==============================")
        print(f"图片文件夹: {images_dir}")
        if csv_file:
            print(f"CSV 文件:   {csv_file}")
        print("==============================\n")

        # 启动后台线程执行转换
        threading.Thread(target=self.process_task, args=(images_dir, csv_file), daemon=True).start()

    def process_task(self, images_dir, csv_file):
        try:
            # 1. 检测是否需要注册 HEIF
            current_exts = set(self.CONVERT_EXTS)
            if any(glob.glob(os.path.join(images_dir, f"*{ext}")) for ext in self.HEIF_EXTS):
                if self._register_heif():
                    current_exts |= self.HEIF_EXTS
                else:
                    print("  [跳过] 检测到 .heic/.heif 文件，但 pillow-heif 不可用，将不处理。")

            # 2. 收集文件 (修复重复处理 bug)
            raw_files = []
            for ext in current_exts:
                raw_files.extend(glob.glob(os.path.join(images_dir, f"*{ext}")))
                raw_files.extend(glob.glob(os.path.join(images_dir, f"*{ext.upper()}")))

            # 针对 Windows 不区分大小写的特性，使用字典进行严格去重
            unique_files_dict = {}
            for f in raw_files:
                unique_files_dict[os.path.normcase(f)] = f
            convert_files = list(unique_files_dict.values())

            total = len(convert_files)
            if total == 0:
                print(f"未找到需要转换的图片（支持: {', '.join(sorted(current_exts))}）")
                self.reset_ui()
                return

            ext_count = Counter(os.path.splitext(f)[1].lower() for f in convert_files)
            print(f"实际需要转换 {total} 张图片 (已自动去重):")
            for ext, count in sorted(ext_count.items()):
                print(f"  {ext}: {count} 张")
            print(f"\n开始转换...\n")

            success = 0
            fail = 0

            # 3. 转换图片
            for i, src_path in enumerate(convert_files, 1):
                jpg_path = os.path.splitext(src_path)[0] + ".jpg"
                try:
                    img = Image.open(src_path)
                    if img.mode in ("RGBA", "LA", "P"):
                        background = Image.new("RGB", img.size, (255, 255, 255))
                        if img.mode == "P":
                            img = img.convert("RGBA")
                        background.paste(img, mask=img.split()[-1] if "A" in img.mode else None)
                        img = background
                    elif img.mode != "RGB":
                        img = img.convert("RGB")

                    img.save(jpg_path, "JPEG", quality=90)
                    img.close()
                    os.remove(src_path)
                    success += 1
                except Exception as e:
                    print(f"  [失败] {os.path.basename(src_path)}: {e}")
                    fail += 1

                # 更新进度条和日志
                self.progress_var.set((i / total) * 100)
                if i % 10 == 0 or i == total:
                    print(f"  进度: {i}/{total}  (成功 {success}, 失败 {fail})")

            print(f"\n图片转换完成！成功 {success}，失败 {fail}\n")

            # 4. 更新 CSV
            if not csv_file:
                print("未提供 CSV 文件，跳过更新。")
            elif not os.path.isfile(csv_file):
                print(f"[警告] CSV 文件不存在: {csv_file}")
            else:
                print("正在更新 CSV 文件...")
                self._update_csv(csv_file, current_exts)

            print("\n✅ 所有的任务已完成！")

        except Exception as e:
            print(f"\n[严重错误] 程序运行中发生异常: {e}")
        finally:
            self.reset_ui()

    def _update_csv(self, csv_file, current_exts):
        try:
            with open(csv_file, "r", encoding="utf-8") as f:
                reader = csv.reader(f)
                rows = list(reader)

            if not rows:
                print("CSV 文件为空。")
                return

            header = rows[0]
            img_col = None
            for idx, col_name in enumerate(header):
                if col_name.strip().lower() == "images":
                    img_col = idx
                    break

            if img_col is None:
                print("[警告] CSV 中未找到 'Images' 列，跳过 CSV 更新。")
                return

            updated = 0
            for row in rows[1:]:
                if img_col < len(row):
                    original = row[img_col]
                    new_val = original
                    for ext in current_exts:
                        # 替换小写和大写后缀
                        new_val = new_val.replace(ext, ".jpg").replace(ext.upper(), ".jpg")
                    if new_val != original:
                        row[img_col] = new_val
                        updated += 1

            backup_path = csv_file + ".bak"
            if os.path.exists(backup_path):
                os.remove(backup_path)  # 如果之前的备份存在则移除
            os.rename(csv_file, backup_path)

            with open(csv_file, "w", encoding="utf-8", newline="") as f:
                writer = csv.writer(f)
                writer.writerows(rows)

            print(f"CSV 已更新 {updated} 行，原文件已备份为: {os.path.basename(backup_path)}")

        except Exception as e:
            print(f"[CSV 错误] 更新 CSV 文件时出错: {e}")

    def reset_ui(self):
        # 使用 after 确保在主线程恢复按钮状态
        self.after(0, lambda: self.btn_start.config(state=tk.NORMAL))


if __name__ == "__main__":
    app = ImageConverterApp()
    app.mainloop()