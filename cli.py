"""docgrep CLI 本体。"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path
from typing import Any, Dict, List, Optional

from config import ConfigError, load_config
from extractors import ExtractorRegistry
from priority import apply_priority
from search import FileResult, Searcher
from selfcheck import print_result, run_selfcheck
from utils import apply_timestamp, ensure_dir, setup_logging, timestamp_slug
from walker import iter_files


# docgrep.py/cli.py が置かれているディレクトリ。フルパスで起動された場合の
# 既定 config.yaml 探索の基準にする。
SCRIPT_DIR = Path(__file__).resolve().parent

# Exit code（grep 互換 + 設定エラー + 中断）
EXIT_HITS_FOUND = 0
EXIT_NO_HITS = 1
EXIT_CONFIG_ERROR = 2
EXIT_INTERRUPTED = 130


def _resolve_default_config() -> Optional[str]:
    """`-c` 省略時の config.yaml を CWD → SCRIPT_DIR の順で探す。

    どちらにも無ければ None（既定値だけで動作）。
    """
    cwd_cand = Path.cwd() / "config.yaml"
    if cwd_cand.is_file():
        return str(cwd_cand)
    script_cand = SCRIPT_DIR / "config.yaml"
    if script_cand.is_file():
        return str(script_cand)
    return None


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="docgrep",
        description="ファイルサーバ全文検索ツール（Office/テキスト対応・逐次走査）",
    )
    p.add_argument("keywords", nargs="*", help="検索キーワード（複数指定可）")
    p.add_argument("-c", "--config", default=None,
                   help="設定ファイル（既定: CWD の config.yaml → スクリプト同梱の config.yaml の順で自動検出）")
    p.add_argument("-p", "--path", action="append",
                   help="検索対象パス（複数指定可、設定の paths を上書き）")
    p.add_argument("--mode", choices=["keyword", "regex", "fuzzy"], help="検索モード")
    p.add_argument("--operator", choices=["and", "or"], help="複数キーワード時の演算子")
    p.add_argument("--case-sensitive", action="store_true", help="大文字小文字を区別する")
    p.add_argument("--no-normalize-width", action="store_true",
                   help="全角/半角の NFKC 正規化を無効化")
    p.add_argument("--fuzzy-threshold", type=float, help="あいまい検索のしきい値 (0.0-1.0)")
    p.add_argument("--snippet-chars", type=int, help="ヒット箇所前後の文字数")
    p.add_argument("--excel", help="Excel 出力先パス（{ts} はタイムスタンプに置換）")
    p.add_argument("--html", help="HTML 出力先パス（{ts} はタイムスタンプに置換）")
    p.add_argument("--no-console", action="store_true", help="コンソール出力を抑制")
    p.add_argument("--no-office-check", action="store_true",
                   help="MS Office チェックをスキップ（runtime.require_office=false 相当）")
    p.add_argument("-v", "--verbose", action="store_true", help="DEBUG レベルログを出力")
    p.add_argument("--quiet", action="store_true", help="進捗バー等を抑制（ログは WARNING 以上）")
    return p


def _merge_cli(cfg: Dict[str, Any], args: argparse.Namespace) -> Dict[str, Any]:
    if args.path:
        cfg["paths"] = args.path
    if args.mode:
        cfg["search"]["mode"] = args.mode
    if args.operator:
        cfg["search"]["operator"] = args.operator
    if args.case_sensitive:
        cfg["search"]["case_sensitive"] = True
    if args.no_normalize_width:
        cfg["search"]["normalize_width"] = False
    if args.fuzzy_threshold is not None:
        cfg["search"]["fuzzy_threshold"] = args.fuzzy_threshold
    if args.snippet_chars is not None:
        cfg["search"]["snippet_chars"] = args.snippet_chars
    if args.keywords:
        cfg["search"]["keywords"] = args.keywords
    if args.excel:
        cfg["output"]["excel"]["enabled"] = True
        cfg["output"]["excel"]["path"] = args.excel
    if args.html:
        cfg["output"]["html"]["enabled"] = True
        cfg["output"]["html"]["path"] = args.html
    if args.no_console:
        cfg["output"]["console"] = False
    if args.no_office_check:
        cfg["runtime"]["require_office"] = False
    return cfg


def main(argv: Optional[List[str]] = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    log = setup_logging(verbose=args.verbose, quiet=args.quiet)

    # === 設定読込 ===
    config_path = args.config if args.config else _resolve_default_config()
    if config_path:
        log.info("設定ファイル: %s", config_path)
    else:
        log.info("設定ファイル: なし（既定値で動作）")
    try:
        cfg = load_config(config_path, default_base_dir=SCRIPT_DIR)
    except ConfigError as e:
        log.error("設定エラー: %s", e)
        return EXIT_CONFIG_ERROR
    except Exception as e:
        log.error("設定ファイル読込で予期せぬエラー: %s", e)
        return EXIT_CONFIG_ERROR
    cfg = _merge_cli(cfg, args)

    # === セルフチェック ===
    check = run_selfcheck(require_office=cfg["runtime"]["require_office"])
    print_result(check)
    if not check.ok:
        return EXIT_CONFIG_ERROR

    # === キーワード必須 ===
    keywords: List[str] = list(cfg["search"]["keywords"] or [])
    if not keywords:
        log.error("キーワードを指定してください（CLI 引数または config の search.keywords）")
        return EXIT_CONFIG_ERROR

    apply_priority(cfg["runtime"]["process_priority"])

    # === Searcher ===
    try:
        searcher = Searcher(
            mode=cfg["search"]["mode"],
            keywords=keywords,
            operator=cfg["search"]["operator"],
            case_sensitive=cfg["search"]["case_sensitive"],
            normalize_width=cfg["search"]["normalize_width"],
            fuzzy_threshold=cfg["search"]["fuzzy_threshold"],
            snippet_chars=cfg["search"]["snippet_chars"],
            max_hits_per_file=cfg["search"]["max_hits_per_file"],
        )
    except Exception as e:
        log.error("検索条件の構築に失敗: %s", e)
        return EXIT_CONFIG_ERROR

    # === COM 抽出器（Windows + MS Office があれば有効） ===
    com = None
    try:
        from extractors.office_com import OfficeCom, is_available
        if is_available():
            com = OfficeCom(recycle_every=cfg["runtime"]["com_recycle_every"])
    except Exception as e:
        log.warning("COM 抽出器の初期化に失敗: %s", e)

    registry = ExtractorRegistry(com=com)

    # === 走査対象: OneNote エクスポートフォルダを自動追加 ===
    paths: List[str] = list(cfg["paths"])
    onenote_dir = cfg.get("onenote_export_dir")
    if onenote_dir and os.path.isdir(onenote_dir) and onenote_dir not in paths:
        paths.append(onenote_dir)

    file_list = list(iter_files(
        paths=paths,
        exclude_dirs=cfg["exclude"]["dirs"],
        exclude_patterns=cfg["exclude"]["patterns"],
        extensions=cfg["extensions"],
        max_size_mb=cfg["exclude"]["max_file_size_mb"],
    ))
    total = len(file_list)
    if total == 0:
        log.info("走査対象のファイルが見つかりませんでした。paths と extensions を確認してください。")
        if com is not None:
            com.shutdown()
        return EXIT_NO_HITS

    try:
        from tqdm import tqdm  # type: ignore
    except ImportError:
        tqdm = lambda it, **kw: it  # type: ignore

    log.info("走査開始: %d ファイル / paths=%s", total, paths)
    interrupted = False
    hit_results: List[FileResult] = []
    error_results: List[FileResult] = []
    skipped = 0
    hits_total = 0
    skip_breakdown: Dict[str, int] = {}
    start = time.monotonic()

    pbar = tqdm(file_list, desc="scan", unit="file", disable=args.quiet)
    try:
        for path in pbar:
            try:
                segments, skip_reason = registry.extract(path)
            except KeyboardInterrupt:
                raise
            except Exception as e:
                error_results.append(FileResult(path=path, error=f"extract_error: {e}"))
                log.debug("extract_error: %s: %s", path, e)
                continue
            if skip_reason:
                skipped += 1
                skip_breakdown[skip_reason] = skip_breakdown.get(skip_reason, 0) + 1
                continue
            if not segments:
                continue
            try:
                hits = searcher.search(segments)
            except Exception as e:
                error_results.append(FileResult(path=path, error=f"search_error: {e}"))
                log.debug("search_error: %s: %s", path, e)
                continue
            if hits:
                fr = FileResult(path=path, hits=hits)
                hit_results.append(fr)
                hits_total += len(hits)
                if cfg["output"]["console"]:
                    from reporter.console import render_console
                    render_console(fr, quiet=args.quiet)
    except KeyboardInterrupt:
        interrupted = True
        log.warning("中断されました。これまでの結果を出力します。")
    finally:
        if com is not None:
            com.shutdown()

    elapsed = time.monotonic() - start

    summary: Dict[str, Any] = {
        "走査ファイル数": total,
        "ヒットファイル数": len(hit_results),
        "ヒット件数合計": hits_total,
        "スキップ数": skipped,
        "エラー数": len(error_results),
        "処理時間(秒)": f"{elapsed:.2f}",
        "検索モード": cfg["search"]["mode"],
        "演算子": cfg["search"]["operator"],
        "キーワード": ", ".join(keywords),
    }
    if skip_breakdown:
        summary["スキップ内訳"] = ", ".join(f"{k}={v}" for k, v in skip_breakdown.items())
    if interrupted:
        summary["中断"] = "Ctrl+C により途中終了"

    # === 出力 ===
    slug = timestamp_slug()
    if cfg["output"]["excel"]["enabled"]:
        try:
            from reporter.excel import write_excel
            excel_path = apply_timestamp(cfg["output"]["excel"]["path"], slug)
            ensure_dir(Path(excel_path).resolve().parent)
            write_excel(excel_path, hit_results, summary, errors=error_results)
            log.info("Excel 出力: %s", excel_path)
        except Exception as e:
            log.warning("Excel 出力に失敗: %s", e)

    if cfg["output"]["html"]["enabled"]:
        try:
            from reporter.html import write_html
            html_path = apply_timestamp(cfg["output"]["html"]["path"], slug)
            ensure_dir(Path(html_path).resolve().parent)
            write_html(html_path, hit_results, summary, searcher, errors=error_results)
            log.info("HTML 出力: %s", html_path)
        except Exception as e:
            log.warning("HTML 出力に失敗: %s", e)

    # === コンソールサマリ（自動化スクリプトからも読めるよう print で出力） ===
    print()
    print("=" * 60)
    print(" 検索結果サマリ")
    print("=" * 60)
    for k, v in summary.items():
        print(f"  {k:<14}: {v}")
    print("=" * 60)

    if interrupted:
        return EXIT_INTERRUPTED
    return EXIT_HITS_FOUND if hits_total > 0 else EXIT_NO_HITS
