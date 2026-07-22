from __future__ import annotations

import shutil
import sys
from pathlib import Path


def resource_root() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent.parent))


def resource_path(name: str) -> Path:
    path = resource_root() / "assets" / name
    if not path.exists():
        project_root = Path(__file__).resolve().parent.parent
        # 销售出库示例使用脱敏数据；模板和 Logo 仍从项目根目录读取。
        for fallback in (project_root / "sample_data" / name, project_root / name):
            if fallback.exists():
                return fallback
    return path


def export_resource(name: str, destination: str | Path) -> None:
    shutil.copyfile(resource_path(name), destination)
