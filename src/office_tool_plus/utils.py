import os
from pathlib import Path


def ensure_unique_path(path):
    """
    确保文件路径唯一，若存在同名文件则在文件名后添加序号
    """
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}_{counter}{ext}"
        counter += 1
    return path


def new_file_path(file_path, suffix, dir_path=None):
    """
    根据原文件路径和新扩展名生成新的文件路径
    file_path：源文件的路径
    suffix：新文件的后缀
    dir_path：新文件存放的路径
    """
    file_path = Path(file_path)
    new_name = f"{file_path.stem}.{suffix}"
    dir_path = dir_path or os.path.dirname(file_path) + os.sep
    if dir_path is not None:
        dir_path = Path(dir_path)
        dir_path.mkdir(parents=True, exist_ok=True)  # 确保输出目录存在
        return ensure_unique_path(str(dir_path / new_name))
    else:
        return ensure_unique_path(str(file_path.with_suffix(f".{suffix}")))
