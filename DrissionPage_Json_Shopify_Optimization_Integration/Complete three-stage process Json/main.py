"""统一三阶段工作台入口。"""
import tkinter as tk
from app.window import Workbench

def main():
    """创建 Tk 根窗口、挂载工作台并进入主事件循环。"""
    root = tk.Tk()
    Workbench(root)
    root.mainloop()

if __name__ == "__main__":
    main()
