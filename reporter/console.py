"""コンソール出力。"""
from __future__ import annotations

from search import FileResult


def render_console(file_result: FileResult, quiet: bool = False) -> None:
    if file_result.error:
        if not quiet:
            print(f"[ERROR] {file_result.path}: {file_result.error}")
        return
    if not file_result.hits:
        return
    print(f"\n{file_result.path}  ({len(file_result.hits)} hits)")
    for hit in file_result.hits:
        prefix = f"[{hit.locator}] " if hit.locator else ""
        print(f"  - {prefix}{hit.snippet}")
