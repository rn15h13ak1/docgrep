"""設定ファイル（YAML）読込と既定値マージ。

PyYAML が利用できない場合は JSON フォールバック（同名 .json）も受け付ける。
"""
from __future__ import annotations

import json
import os
from copy import deepcopy
from pathlib import Path
from typing import Any, Dict, Optional


XLSX_GRANULARITY_CHOICES = ("cell", "row")


DEFAULT_CONFIG: Dict[str, Any] = {
    "paths": ["."],
    "onenote_export_dir": "./onenote_export",
    "extensions": [
        ".txt", ".csv", ".log", ".md", ".tsv",
        ".xlsx", ".xlsm", ".xls",
        ".docx", ".doc",
        ".pptx", ".ppt",
    ],
    # xlsx 抽出時の Segment 粒度: "cell" (既定) または "row"。大規模 xlsx で
    # メモリ/時間を抑えたい場合は "row" にすると Segment 数が列数分減る。
    "xlsx_granularity": "cell",
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
        # 抽出結果の SQLite キャッシュ。enabled=true なら2回目以降の検索が劇的に速くなる。
        # path に "{ts}" や相対パスを使えるが、通常は固定パスにする。
        "cache": {
            "enabled": False,
            "path": "reports/.docgrep_cache.sqlite",
        },
    },
    "output": {
        "console": True,
        # path に "{ts}" を含めると実行時刻 (YYYYMMDD-HHMMSS) に置換される。
        # 既定では reports/ フォルダ配下に履歴付きで出力される。
        "excel": {
            "enabled": True,
            "path": "reports/search_result_{ts}.xlsx",
        },
        "html": {
            "enabled": True,
            "path": "reports/search_result_{ts}.html",
            # latest_path を指定するとタイムスタンプ無しの「最新版」が同じ内容で
            # 上書き出力される。ブックマーク用途やリンク固定に便利。空文字 / null で無効化。
            "latest_path": "reports/search_result_latest.html",
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
    _validate_types(cfg, path or "<defaults>")
    return cfg


def _validate_types(cfg: Dict[str, Any], source: str) -> None:
    """設定値の型・許容値・範囲を検証する。違反は ConfigError でまとめて報告する。"""
    errs: list[str] = []

    def _need(value: Any, expected: type, key: str) -> bool:
        if not isinstance(value, expected):
            errs.append(f"{key}: {type(expected).__name__} を期待しましたが "
                        f"{type(value).__name__} ({value!r}) でした")
            return False
        return True

    s = cfg.get("search", {})
    if _need(s.get("mode"), str, "search.mode") and s["mode"] not in ("keyword", "regex", "fuzzy"):
        errs.append(f"search.mode: 'keyword'/'regex'/'fuzzy' のいずれかを指定 (現在: {s['mode']!r})")
    if _need(s.get("operator"), str, "search.operator") and s["operator"] not in ("and", "or"):
        errs.append(f"search.operator: 'and'/'or' のいずれかを指定 (現在: {s['operator']!r})")
    if _need(s.get("case_sensitive"), bool, "search.case_sensitive"):
        pass
    if _need(s.get("normalize_width"), bool, "search.normalize_width"):
        pass
    th = s.get("fuzzy_threshold")
    if _need(th, (int, float), "search.fuzzy_threshold"):
        if not (0.0 <= float(th) <= 1.0):
            errs.append(f"search.fuzzy_threshold: 0.0〜1.0 の範囲で指定 (現在: {th})")
    if _need(s.get("snippet_chars"), int, "search.snippet_chars") and s["snippet_chars"] < 0:
        errs.append("search.snippet_chars: 0 以上の整数")
    if _need(s.get("max_hits_per_file"), int, "search.max_hits_per_file") and s["max_hits_per_file"] < 1:
        errs.append("search.max_hits_per_file: 1 以上の整数")

    r = cfg.get("runtime", {})
    if _need(r.get("require_office"), bool, "runtime.require_office"):
        pass
    if _need(r.get("parallel"), int, "runtime.parallel") and r["parallel"] < 0:
        errs.append("runtime.parallel: 0 以上の整数（0=auto）")
    pp = r.get("process_priority")
    if _need(pp, str, "runtime.process_priority") and pp not in ("normal", "below_normal", "idle"):
        errs.append(f"runtime.process_priority: 'normal'/'below_normal'/'idle' (現在: {pp!r})")
    if _need(r.get("com_recycle_every"), int, "runtime.com_recycle_every") and r["com_recycle_every"] < 1:
        errs.append("runtime.com_recycle_every: 1 以上の整数")
    pt = r.get("per_file_timeout_sec")
    if _need(pt, (int, float), "runtime.per_file_timeout_sec") and float(pt) < 0:
        errs.append("runtime.per_file_timeout_sec: 0 以上 (0=無効)")
    cache_cfg = r.get("cache")
    if cache_cfg is not None and _need(cache_cfg, dict, "runtime.cache"):
        if "enabled" in cache_cfg:
            _need(cache_cfg["enabled"], bool, "runtime.cache.enabled")

    o = cfg.get("output", {})
    if _need(o.get("console"), bool, "output.console"):
        pass
    for k in ("excel", "html"):
        sect = o.get(k)
        if sect is not None and _need(sect, dict, f"output.{k}"):
            if "enabled" in sect:
                _need(sect["enabled"], bool, f"output.{k}.enabled")

    gran = cfg.get("xlsx_granularity")
    if gran is not None and _need(gran, str, "xlsx_granularity") \
            and gran not in XLSX_GRANULARITY_CHOICES:
        errs.append(f"xlsx_granularity: {XLSX_GRANULARITY_CHOICES} のいずれか (現在: {gran!r})")

    ext = cfg.get("extensions")
    if ext is not None:
        if not isinstance(ext, list):
            errs.append(f"extensions: list を期待しましたが {type(ext).__name__}")
        else:
            for i, v in enumerate(ext):
                if not isinstance(v, str):
                    errs.append(f"extensions[{i}]: 文字列を期待しましたが {type(v).__name__}")

    excl = cfg.get("exclude")
    if isinstance(excl, dict):
        mfs = excl.get("max_file_size_mb")
        if mfs is not None and _need(mfs, (int, float), "exclude.max_file_size_mb") \
                and float(mfs) < 0:
            errs.append("exclude.max_file_size_mb: 0 以上 (0=無効)")

    if errs:
        raise ConfigError(
            f"設定ファイル {source} の値検証に失敗:\n  - " + "\n  - ".join(errs)
        )


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

    # runtime.cache.path
    runtime = cfg.get("runtime")
    if isinstance(runtime, dict):
        cache_cfg = runtime.get("cache")
        if isinstance(cache_cfg, dict) and isinstance(cache_cfg.get("path"), str):
            cache_cfg["path"] = _r(cache_cfg["path"])


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

    runtime = data.get("runtime")
    if isinstance(runtime, dict):
        cache_cfg = runtime.get("cache")
        if isinstance(cache_cfg, dict):
            cp = cache_cfg.get("path")
            if isinstance(cp, str) and "\\" in cp:
                violations.append(f"runtime.cache.path = {cp!r}")

    if violations:
        raise ConfigError(
            f"設定ファイル {source} のパス値にバックスラッシュ (\\) が含まれています。"
            "フォワードスラッシュ (/) のみ使用可能です"
            "（UNC は //server/share/...、ドライブは Z:/path のように記述してください）:\n  - "
            + "\n  - ".join(violations)
        )
