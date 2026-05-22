"""ファイル走査（os.walk + 拡張子フィルタ + 除外）。"""
from __future__ import annotations

import fnmatch
import os
from typing import Iterable, Iterator, Sequence


def iter_files(
    paths: Sequence[str],
    exclude_dirs: Iterable[str],
    exclude_patterns: Iterable[str],
    extensions: Iterable[str],
    max_size_mb: int | None,
) -> Iterator[str]:
    """対象パス配下を再帰走査して条件に合うファイルを yield する。

    extensions に "*" を含めるか空リストを渡すと **拡張子フィルタを無効化** する。
    その場合は全ファイルを走査し、抽出ディスパッチでバイナリ判定が行われる。
    """
    ext_list = [e.lower() for e in extensions if e]
    wildcard = (not ext_list) or ("*" in ext_list)
    ext_set: set[str] = set() if wildcard else {e for e in ext_list if e != "*"}
    exc_dirs_set = set(exclude_dirs)
    exc_pat_list = list(exclude_patterns)
    max_bytes = max_size_mb * 1024 * 1024 if max_size_mb else None

    for root_path in paths:
        if not root_path:
            continue
        if not os.path.exists(root_path):
            continue
        # ファイル単体指定にも対応
        if os.path.isfile(root_path):
            if _matches_filters(root_path, ext_set, exc_pat_list, max_bytes):
                yield root_path
            continue

        for dirpath, dirnames, filenames in os.walk(root_path):
            # 除外フォルダ（in-place で dirnames を削ると os.walk が降りなくなる）
            dirnames[:] = [d for d in dirnames if d not in exc_dirs_set]
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                if _matches_filters(full, ext_set, exc_pat_list, max_bytes):
                    yield full


def _matches_filters(
    full: str,
    ext_set: set[str],
    exc_pat_list: list[str],
    max_bytes: int | None,
) -> bool:
    fn = os.path.basename(full)
    # 除外パターン
    for pat in exc_pat_list:
        if fnmatch.fnmatch(fn, pat):
            return False
    # 拡張子フィルタ（ext_set が空ならフィルタしない＝全拡張子許可）
    if ext_set:
        ext = os.path.splitext(fn)[1].lower()
        if ext not in ext_set:
            return False
    # サイズ上限
    if max_bytes is not None:
        try:
            if os.path.getsize(full) > max_bytes:
                return False
        except OSError:
            return False
    return True
