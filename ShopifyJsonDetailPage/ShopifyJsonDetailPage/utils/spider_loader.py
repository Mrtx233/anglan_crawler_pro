"""
自定义 Spider Loader，支持任意名称的子文件夹。

Scrapy 默认的 walk_modules_iter 要求子文件夹是合法 Python 标识符，
本 loader 绕过此限制，从文件路径直接加载 spider 模块，
使 spiders/midsummera.com男女装-日常休闲款/ 等文件夹也能被发现。
"""

import importlib
import importlib.util
import os

from scrapy.spiderloader import SpiderLoader
from scrapy.utils.misc import walk_modules_iter


class RecursiveSpiderLoader(SpiderLoader):

    def _load_all_spiders(self):
        for module_name in self.spider_modules:
            try:
                pkg = importlib.import_module(module_name)
            except ImportError:
                continue

            pkg_dir = os.path.dirname(pkg.__file__)

            # 顶层 spider 文件
            for entry in os.listdir(pkg_dir):
                if entry.endswith(".py") and not entry.startswith("_"):
                    self._load_from_file(
                        os.path.join(pkg_dir, entry),
                        f"{module_name}.{entry[:-3]}",
                    )

            # 子文件夹（允许任意名称）
            for entry in sorted(os.listdir(pkg_dir)):
                subdir = os.path.join(pkg_dir, entry)
                if not os.path.isdir(subdir) or entry.startswith((".", "__")):
                    continue
                for py_file in sorted(os.listdir(subdir)):
                    if py_file.endswith(".py") and not py_file.startswith("_"):
                        mod_name = f"{module_name}._sub_.{entry}.{py_file[:-3]}"
                        self._load_from_file(
                            os.path.join(subdir, py_file),
                            mod_name,
                        )

        self._check_name_duplicates()

    def _load_from_file(self, filepath, module_name):
        try:
            spec = importlib.util.spec_from_file_location(module_name, filepath)
            if spec is None or spec.loader is None:
                return
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            self._load_spiders(mod)
        except Exception as e:
            if self.warn_only:
                import warnings
                warnings.warn(
                    f"Could not load spider from {filepath}: {e}",
                    stacklevel=2,
                    category=RuntimeWarning,
                )
