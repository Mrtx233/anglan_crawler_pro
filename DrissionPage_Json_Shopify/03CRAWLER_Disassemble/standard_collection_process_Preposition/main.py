"""Run with python main.py or python -m standard_collection_process_Preposition.main."""

from pathlib import Path
import sys

if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from standard_collection_process_Preposition.ui.app import ShopifyPipelineApp
else:
    from .ui.app import ShopifyPipelineApp
import tkinter as tk


def main():
    root = tk.Tk()
    ShopifyPipelineApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
