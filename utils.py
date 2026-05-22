"""ロギングと共通ユーティリティ。"""
from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional


_LOGGER_NAME = "docgrep"


def setup_logging(verbose: bool = False, quiet: bool = False) -> logging.Logger:
    """docgrep 用ロガーを初期化する。

    verbose=True で DEBUG、quiet=True で WARNING、既定は INFO。
    複数回呼ばれてもハンドラの重複登録は行わない。
    """
    logger = logging.getLogger(_LOGGER_NAME)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(
            logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
        )
        logger.addHandler(handler)
        logger.propagate = False
    if verbose:
        logger.setLevel(logging.DEBUG)
    elif quiet:
        logger.setLevel(logging.WARNING)
    else:
        logger.setLevel(logging.INFO)
    return logger


def get_logger() -> logging.Logger:
    return logging.getLogger(_LOGGER_NAME)


def timestamp_slug(dt: Optional[datetime] = None) -> str:
    """出力ファイル名向けのタイムスタンプ (YYYYMMDD-HHMMSS)。"""
    return (dt or datetime.now()).strftime("%Y%m%d-%H%M%S")


def apply_timestamp(path_str: str, slug: str) -> str:
    """パス文字列に含まれる {ts} を slug で置換する。

    {ts} を含まない場合は元の文字列をそのまま返す（後方互換）。
    """
    if not path_str:
        return path_str
    return path_str.replace("{ts}", slug)


def ensure_dir(path: Path) -> Path:
    """親ディレクトリを含めて作成する（既存なら何もしない）。"""
    path.mkdir(parents=True, exist_ok=True)
    return path


def display_path(path: str) -> str:
    """表示用にパスを `/` 区切りで統一する（OS によらず一貫表示）。"""
    if not path:
        return path
    return path.replace("\\", "/")
