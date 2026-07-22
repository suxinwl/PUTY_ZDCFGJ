from __future__ import annotations

import shutil
import subprocess
import sys
import tkinter as tk
import os
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parent
ASSETS = ROOT / "assets"
DIST = ROOT / "dist"


def prepare_assets() -> Path:
    ASSETS.mkdir(exist_ok=True)
    sources = {
        "logo.png": ROOT / "logo.png",
        "202601销售出库.xlsx": ROOT / "sample_data" / "202601销售出库.xlsx",
        "表头表尾格式.xlsx": ROOT / "表头表尾格式.xlsx",
    }
    for name, source in sources.items():
        if not source.exists():
            raise FileNotFoundError(f"缺少打包资源：{source}")
        shutil.copyfile(source, ASSETS / name)
    icon_path = ASSETS / "app.ico"
    image = Image.open(ROOT / "logo.png").convert("RGBA")
    image.save(icon_path, format="ICO", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    return icon_path


def verify_tkinter() -> None:
    """避免 PyInstaller 静默生成一个不包含 GUI 运行库的无效程序。"""
    try:
        root = tk.Tk()
        root.withdraw()
        root.update_idletasks()
        root.destroy()
    except Exception as exc:
        raise RuntimeError("当前 Python 的 Tkinter/Tcl-Tk 不完整，请改用包含 Tcl/Tk 的 Python 后重新打包。") from exc


def main() -> None:
    verify_tkinter()
    icon = prepare_assets()
    executable_name = os.environ.get("BILL_SPLITTER_EXE_NAME", "账单拆分插入工具")
    separator = ";" if sys.platform == "win32" else ":"
    command = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--windowed",
        "--name", executable_name,
        "--icon", str(icon),
        "--add-data", f"{ASSETS}{separator}assets",
        # openpyxl 会探测这些可选数据科学组件；本工具不使用它们，排除后体积更小且避免误收集构建机模块。
        "--exclude-module", "numpy",
        "--exclude-module", "pandas",
        "--exclude-module", "pyarrow",
        "--exclude-module", "lxml",
        "app.py",
    ]
    subprocess.run(command, cwd=ROOT, check=True)
    executable = DIST / (f"{executable_name}.exe" if sys.platform == "win32" else executable_name)
    print(f"构建完成：{executable}")


if __name__ == "__main__":
    main()
