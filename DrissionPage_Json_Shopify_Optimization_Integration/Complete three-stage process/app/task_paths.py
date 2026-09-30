"""固定输入根目录下的批次、负责人和任务目录。"""
from pathlib import Path
import unicodedata

INPUT_ROOT = Path(__file__).resolve().parents[2] / '01INPUT_XLSX'


def directories(parent):
    parent = Path(parent)
    if not parent.is_dir():
        return []
    return sorted(p.name for p in parent.iterdir()
                  if p.is_dir() and not p.is_symlink() and not p.name.startswith('.'))


def clean_task_name(value):
    value = unicodedata.normalize('NFKC', value)
    name = ''.join(c for c in value if not c.isspace()
                   and not unicodedata.category(c).startswith('C')
                   and c not in '<>:"/\\|?*').strip('.')
    if not name:
        raise ValueError('清理后的任务文件夹名称不能为空')
    if len(name.encode('utf-8')) > 255:
        raise ValueError('任务文件夹名称过长')
    return name


def task_path(batch, owner, name):
    parts = []
    parent = INPUT_ROOT
    for label, value in (("批次", batch), ("负责人", owner), ("任务文件夹", name)):
        if not value.strip():
            raise ValueError(f'请输入{label}')
        # 已有目录沿用原名；新名称统一清理。
        part = value if value in directories(parent) else clean_task_name(value)
        parent = parent / part
        if parent.is_symlink() or (parent.exists() and not parent.is_dir()):
            raise ValueError(f'{label}路径已被文件或符号链接占用')
        parts.append(part)
    return INPUT_ROOT.joinpath(*parts)
