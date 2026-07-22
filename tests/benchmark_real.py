"""对工作区内真实样例做单文件分阶段计时（不属于单元测试）。"""
from __future__ import annotations

import tempfile
import time
from pathlib import Path

from bill_splitter.core import BillSplitter, TemplateSnapshot, open_reader


def mark(label: str, started: float) -> float:
    now = time.perf_counter()
    print(f"{label}: {now - started:.3f}s", flush=True)
    return now


files = sorted(Path.cwd().glob("*.xlsx"), key=lambda item: item.stat().st_size)
template_path, source_path = files[0], files[-1]
started = time.perf_counter()
source = open_reader(source_path)
started = mark("读取源文件", started)
splitter = BillSplitter(source_path, template_path, Path.cwd())
groups, _ = splitter._group_rows(source)
started = mark(f"分组（{len(groups)}组）", started)
template = TemplateSnapshot(template_path)
started = mark("读取模板", started)
with tempfile.TemporaryDirectory(prefix="bill_benchmark_", dir=Path.cwd()) as temp_dir:
    splitter.output_dir = Path(temp_dir)
    splitter._write_group(source, template, groups[0], Path(temp_dir) / "one.xlsx")
    mark(f"生成首个账单（{len(groups[0].rows)}行）", started)
source.close()
