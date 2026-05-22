"""設定ファイル（YAML）読込と既定値マージ。

PyYAML が利用できない場合は JSON フォールバック（同名 .json）も受け付ける。
"""
from __future__ import annotations

import json
import os
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Optional


DEFAULT_CONFIG: Dict[str, Any] = {
    "paths": ["."],
    "onenote_export_dir": "./onenote_export",
    "extensions": [
        ".txt", ".csv", ".log", ".md", ".tsv",
        ".xlsx", ".xlsm", ".xls",
        ".docx", ".doc",
        ".pptx", ".ppt",
    ],
    "exclude": {
        "dirs": [".git", ".svn", "node_modules", "__pycache__", ".venv", "venv"],
        "patterns": ["~$*", "*.tmp", ".DS_Store", "Thumbs.db", "_docgrep_meta.json"],
        "max_file_size_mb": 100,
    },
    "search": {
        "mode": "keyword",        # keyword | regex | fuzzy
        "keywords": [],
        "operator": "and",         # and | or
        "case_sensitive": False,
        "normalize_width": True,
        "fuzzy_threshold": 0.80,
        "snippet_chars": 60,
        "max_hits_per_file": 20,
    },
    "runtime": {
        "require_office": True,
        # 並列度: 0=auto(CPU/2), 1=直列, N=N スレッド。COM (Word/PPT/旧Excel) は
        # スレッドセーフではないため常に直列で別ワーカー実行され、ここでの並列度は
        # text/xlsx の抽出にのみ適用される。
        "parallel": 0,
        "process_priority": "below_normal",  # normal | below_normal | idle
        "copy_to_temp": False,
        "com_recycle_every": 30,
        "per_file_timeout_sec": 0,
    },
    "output": {
        "console": True,
        # path に "{ts}" を含めると実行時刻 (YYYYMMDD-HHMMSS) に置換される。
        # 既定では out/ フォルダ配下に履歴付きで出力される。
        "excel": {
            "enabled": True,
            "path": "out/search_result_{ts}.xlsx",
        },
        "html": {
            "enabled": True,
            "path": "out/search_result_{ts}.html",
            # latest_path を指定するとタイムスタンプ無しの「最新版」が同じ内容で
            # 上書き出力される。ブックマーク用途やリンク固定に便利。空文字 / null で無効化。
            "latest_path": "out/search_result_latest.html",
        },
    },
}


class ConfigError(ValueError):
    """設定ファイル起因のエラー。CLI 側で個別ハンドリングする。"""


def _deep_merge(base: Dict[str, Any], override: Dict[str, Any]) -> None:
    for k, v in override.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v


def load_config(
    path: Optional[str] = None,
    default_base_dir: Optional[Path] = None,
) -> Dict[str, Any]:
    """設定ファイルを読み込み既定値とマージして返す。

    path が None または存在しない場合は既定値だけを返す。

    YAML/JSON 内の相対パス（paths / onenote_export_dir / output.*.path）は
    「設定ファイルのある親ディレクトリ」を基準に絶対パス化する。設定ファイル
    が無い場合（既定値のみ）は default_base_dir を基準にする。これにより
    docgrep.py をフルパスで起動しても CWD に依存せず動作する。
    """
    cfg = deepcopy(DEFAULT_CONFIG)
    base_dir: Path

    if path and os.path.exists(path):
        ext = os.path.splitext(path)[1].lower()
        try:
            with open(path, "r", encoding="utf-8") as f:
                data_str = f.read()
        except OSError as e:
            raise ConfigError(f"設定ファイル {path} を開けません: {e}") from e

        if ext == ".json":
            try:
                data = json.loads(data_str) or {}
            except json.JSONDecodeError as e:
                raise ConfigError(f"JSON 解析エラー ({path}): {e}") from e
        else:
            try:
                import yaml  # type: ignore
            except ImportError as e:
                raise ConfigError(
                    "PyYAML がインポートできません。YAML 設定を使うには Anaconda の PyYAML が必要です。"
                ) from e
            try:
                data = yaml.safe_load(data_str) or {}
            except yaml.YAMLError as e:
                raise ConfigError(f"YAML 解析エラー ({path}): {e}") from e

        if not isinstance(data, dict):
            raise ConfigError(f"設定ファイル {path} のトップレベルは辞書である必要があります。")
        _validate_no_backslash(data, path)
        _deep_merge(cfg, data)
        base_dir = Path(path).resolve().parent
    else:
        base_dir = Path(default_base_dir) if default_base_dir else Path.cwd()

    _resolve_paths(cfg, base_dir)
    return cfg


def _resolve_paths(cfg: Dict[str, Any], base_dir: Path) -> None:
    """相対パス値を base_dir 基準で絶対パス化する。

    対象: paths[*] / onenote_export_dir / output.excel.path / output.html.path
    UNC ("//server/share/...") とドライブ指定 ("Z:/...") は is_absolute() で
    True を返すのでそのまま使用する。
    """
    def _r(s: Any) -> Any:
        if not isinstance(s, str) or not s:
            return s
        p = Path(s)
        if p.is_absolute():
            return str(p)
        return str((base_dir / s).resolve())

    paths = cfg.get("paths")
    if isinstance(paths, list):
        cfg["paths"] = [_r(p) for p in paths]

    if isinstance(cfg.get("onenote_export_dir"), str):
        cfg["onenote_export_dir"] = _r(cfg["onenote_export_dir"])

    output = cfg.get("output")
    if isinstance(output, dict):
        for key in ("excel", "html"):
            section = output.get(key)
            if isinstance(section, dict):
                if isinstance(section.get("path"), str):
                    section["path"] = _r(section["path"])
                if isinstance(section.get("latest_path"), str):
                    section["latest_path"] = _r(section["latest_path"])


def _validate_no_backslash(data: Dict[str, Any], source: str) -> None:
    r"""YAML/JSON で読み込んだパス値にバックスラッシュが含まれていないか検証する。

    Windows でも Python の os 関数群はフォワードスラッシュを受け付けるため、
    YAML 上のエスケープ事故（バックスラッシュの数が意図と変わる等）と
    表記揺れを防ぐためにフォワードスラッシュ "/" のみを許可する。

    UNC パスは "//server/share/..." 、ドライブ指定は "Z:/projects" の形式で記述する。
    """
    violations: list[str] = []

    paths = data.get("paths")
    if isinstance(paths, list):
        for i, p in enumerate(paths):
            if isinstance(p, str) and "\\" in p:
                violations.append(f"paths[{i}] = {p!r}")

    onenote = data.get("onenote_export_dir")
    if isinstance(onenote, str) and "\\" in onenote:
        violations.append(f"onenote_export_dir = {onenote!r}")

    output = data.get("output")
    if isinstance(output, dict):
        for key in ("excel", "html"):
            section = output.get(key)
            if isinstance(section, dict):
                op = section.get("path")
                if isinstance(op, str) and "\\" in op:
                    violations.append(f"output.{key}.path = {op!r}")
                lp = section.get("latest_path")
                if isinstance(lp, str) and "\\" in lp:
                    violations.append(f"output.{key}.latest_path = {lp!r}")

    if violations:
        raise ConfigError(
            f"設定ファイル {source} のパス値にバックスラッシュ (\\) が含まれています。"
            "フォワードスラッシュ (/) のみ使用可能です"
            "（UNC は //server/share/...、ドライブは Z:/path のように記述してください）:\n  - "
            + "\n  - ".join(violations)
        )
