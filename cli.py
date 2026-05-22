"""docgrep CLI 本体。

main() は薄いオーケストレーターで、責務ごとに以下のヘルパ関数に分割している:
  _init_runtime    — selfcheck / priority / Searcher 構築
  _init_extractors — COM / parallel-registry / serial-registry / SQLite キャッシュ
  _collect_files   — OneNote ディレクトリ追加 + iter_files
  _run_scan        — 並列+直列の走査ループ
  _build_summary   — サマリ dict 組み立て
  _emit_reports    — Excel/HTML 出力 (latest_path コピー含む)

これにより各段階がユニットテスト可能になっている。
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import logging
import os
import sys
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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

# COM が必要な拡張子（スレッドセーフではないため直列処理）
_COM_EXTS = {".doc", ".docx", ".ppt", ".pptx", ".xls"}


class _NoopBar:
    """tqdm が import できない環境用の no-op 進捗バー。"""

    def __init__(self, *a, **kw) -> None:
        pass

    def update(self, n: int = 1) -> None:
        pass

    def close(self) -> None:
        pass


# =============================================================================
# 1 ファイル単位の抽出 + 検索
# =============================================================================

def _process_file(path: str, registry: ExtractorRegistry, searcher: Searcher,
                  cache=None
                  ) -> Tuple[str, Optional[List], Optional[str], Optional[str]]:
    """1 ファイルを抽出 + 検索する。スレッド安全であること（registry/searcher の状態を変更しない）。

    cache が指定されていれば、mtime/size が一致した場合は抽出をスキップして
    キャッシュから Segment を読む。新規抽出時は結果を cache.put で永続化する。

    Returns:
        (path, hits or None, skip_reason or None, error_message or None)
    """
    segments = None
    if cache is not None:
        try:
            segments = cache.get(path)
        except Exception:
            segments = None

    if segments is None:
        try:
            segments, skip_reason = registry.extract(path)
        except Exception as e:
            return path, None, None, f"extract_error: {e}"
        if skip_reason:
            return path, None, skip_reason, None
        if not segments:
            return path, None, None, None
        if cache is not None:
            try:
                cache.put(path, segments)
            except Exception:
                pass  # キャッシュ書き込み失敗は致命ではない

    try:
        hits = searcher.search(segments)
    except Exception as e:
        return path, None, None, f"search_error: {e}"
    return path, hits, None, None


# =============================================================================
# 設定ファイル探索 + CLI 引数
# =============================================================================

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

    # --- 設定 / パス ---
    g_cfg = p.add_argument_group("設定・対象パス")
    g_cfg.add_argument("-c", "--config", default=None,
                       help="設定ファイル（既定: CWD の config.yaml → スクリプト同梱の config.yaml の順で自動検出）")
    g_cfg.add_argument("-p", "--path", action="append",
                       help="検索対象パス（複数指定可、設定の paths を上書き）")

    # --- 検索条件 ---
    g_search = p.add_argument_group("検索条件")
    g_search.add_argument("--mode", choices=["keyword", "regex", "fuzzy"], help="検索モード")
    g_search.add_argument("--operator", choices=["and", "or"], help="複数キーワード時の演算子")
    g_search.add_argument("--case-sensitive", action="store_true", help="大文字小文字を区別する")
    g_search.add_argument("--no-normalize-width", action="store_true",
                          help="全角/半角の NFKC 正規化を無効化")
    g_search.add_argument("--fuzzy-threshold", type=float,
                          help="あいまい検索のしきい値 (0.0-1.0)")
    g_search.add_argument("--snippet-chars", type=int, help="ヒット箇所前後の文字数")

    # --- 出力 ---
    g_out = p.add_argument_group("出力")
    g_out.add_argument("--excel", help="Excel 出力先パス（{ts} はタイムスタンプに置換）")
    g_out.add_argument("--html", help="HTML 出力先パス（{ts} はタイムスタンプに置換）")
    g_out.add_argument("--no-console", action="store_true", help="コンソール出力を抑制")

    # --- 動作制御 ---
    g_ctl = p.add_argument_group("動作制御")
    g_ctl.add_argument("--no-office-check", action="store_true",
                       help="MS Office チェックをスキップ（runtime.require_office=false 相当）")
    g_ctl.add_argument("-v", "--verbose", action="store_true", help="DEBUG レベルログを出力")
    g_ctl.add_argument("--quiet", action="store_true",
                       help="進捗バー等を抑制（ログは WARNING 以上）")
    g_ctl.add_argument("--max-files", type=int, metavar="N",
                       help="N 件のヒットファイルを得たら走査を打ち切る（0 で無効）")
    g_ctl.add_argument("--first-hit-only", action="store_true",
                       help="最初の 1 ヒットで走査を打ち切る (--max-files=1 相当)")
    g_ctl.add_argument("--ordered-output", action="store_true",
                       help="並列処理時もコンソール出力をファイル入力順に揃える")
    g_ctl.add_argument("--dry-run", action="store_true",
                       help="走査対象の集計だけ表示して終了（実走査・抽出・検索は行わない）")

    # --- キャッシュ ---
    g_cache = p.add_argument_group("抽出キャッシュ")
    g_cache.add_argument("--cache", dest="cache", action="store_true", default=None,
                         help="抽出結果の SQLite キャッシュを有効化")
    g_cache.add_argument("--no-cache", dest="cache", action="store_false",
                         help="抽出結果の SQLite キャッシュを無効化")
    g_cache.add_argument("--cache-path", help="キャッシュ DB のパス上書き")
    g_cache.add_argument("--cache-stats", action="store_true",
                         help="キャッシュ DB の統計を表示して終了")
    g_cache.add_argument("--cache-vacuum", action="store_true",
                         help="キャッシュ DB の孤児セグメントを掃除して VACUUM")
    g_cache.add_argument("--cache-clear", action="store_true",
                         help="キャッシュ DB を全消去して終了")

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
    if args.cache is not None:
        cfg["runtime"].setdefault("cache", {})["enabled"] = args.cache
    if args.cache_path:
        cfg["runtime"].setdefault("cache", {})["path"] = args.cache_path
    return cfg


# =============================================================================
# 段階別ヘルパ
# =============================================================================

@dataclass
class ScanContext:
    """走査・出力に共有される状態。"""
    cfg: Dict[str, Any]
    args: argparse.Namespace
    log: logging.Logger
    searcher: Searcher
    keywords: List[str]
    com: Any = None
    parallel_registry: Optional[ExtractorRegistry] = None
    serial_registry: Optional[ExtractorRegistry] = None
    cache: Any = None
    paths: List[str] = field(default_factory=list)


@dataclass
class ScanResult:
    total: int = 0
    hit_results: List[FileResult] = field(default_factory=list)
    error_results: List[FileResult] = field(default_factory=list)
    skip_files: Dict[str, List[str]] = field(default_factory=dict)
    hits_total: int = 0
    skipped: int = 0
    interrupted: bool = False
    stopped: bool = False        # max_files に到達
    stop_max: int = 0            # 打ち切り時の max_files 値
    parallel_workers: int = 0
    parallel_count: int = 0
    serial_count: int = 0
    elapsed: float = 0.0


def _init_runtime(cfg: Dict[str, Any], args: argparse.Namespace, log: logging.Logger
                  ) -> Optional[Tuple[Searcher, List[str]]]:
    """selfcheck → priority 設定 → Searcher 構築。失敗時 None。"""
    check = run_selfcheck(require_office=cfg["runtime"]["require_office"])
    print_result(check)
    if not check.ok:
        return None

    keywords: List[str] = list(cfg["search"]["keywords"] or [])
    if not keywords:
        log.error("キーワードを指定してください（CLI 引数または config の search.keywords）")
        return None

    apply_priority(cfg["runtime"]["process_priority"])

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
        return None
    return searcher, keywords


def _init_extractors(cfg: Dict[str, Any], log: logging.Logger
                     ) -> Tuple[Any, ExtractorRegistry, ExtractorRegistry, Any]:
    """COM / parallel registry / serial registry / cache を初期化する。"""
    com = None
    try:
        from extractors.office_com import OfficeCom, is_available
        if is_available():
            com = OfficeCom(recycle_every=cfg["runtime"]["com_recycle_every"])
    except Exception as e:
        log.warning("COM 抽出器の初期化に失敗: %s", e)

    xlsx_gran = cfg.get("xlsx_granularity", "cell")
    parallel_registry = ExtractorRegistry(com=None, xlsx_granularity=xlsx_gran)
    serial_registry = ExtractorRegistry(com=com, xlsx_granularity=xlsx_gran)

    cache = None
    cache_cfg = cfg["runtime"].get("cache") or {}
    if cache_cfg.get("enabled"):
        try:
            from cache import SegmentCache
            cache = SegmentCache(Path(cache_cfg.get("path", "reports/.docgrep_cache.sqlite")))
            log.info("キャッシュ有効: %s", cache.db_path)
        except Exception as e:
            log.warning("キャッシュ初期化に失敗: %s", e)
            cache = None
    return com, parallel_registry, serial_registry, cache


def _collect_files(cfg: Dict[str, Any]) -> Tuple[List[str], List[str]]:
    """走査対象 paths と全ファイル一覧を返す。"""
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
    return paths, file_list


def _partition_files(file_list: List[str]) -> Tuple[List[str], List[str]]:
    """COM が必要な拡張子と並列可能な拡張子に分ける。"""
    parallel_files: List[str] = []
    serial_files: List[str] = []
    for p in file_list:
        ext = os.path.splitext(p)[1].lower()
        if ext in _COM_EXTS:
            serial_files.append(p)
        else:
            parallel_files.append(p)
    return parallel_files, serial_files


def _parallel_workers(cfg: Dict[str, Any]) -> int:
    configured = int(cfg["runtime"].get("parallel", 0) or 0)
    if configured <= 0:
        return max(1, (os.cpu_count() or 2) // 2)
    return configured


def _dry_run(file_list: List[str], parallel_files: List[str],
             serial_files: List[str], parallel_workers: int,
             paths: List[str]) -> int:
    """--dry-run: 走査せず対象の集計だけ出して終了する。"""
    total = len(file_list)
    print()
    print("=" * 60)
    print(" DRY RUN — 走査対象の集計（抽出・検索は行いません）")
    print("=" * 60)
    print(f"  対象パス      : {paths}")
    print(f"  ファイル総数  : {total}")
    print(f"  並列対象      : {len(parallel_files)}")
    print(f"  COM 直列対象  : {len(serial_files)}")
    print(f"  並列度        : {parallel_workers}")

    # 拡張子別の集計
    ext_count: Dict[str, int] = {}
    for f in file_list:
        e = os.path.splitext(f)[1].lower() or "(なし)"
        ext_count[e] = ext_count.get(e, 0) + 1
    if ext_count:
        print("  拡張子別      :")
        for ext, n in sorted(ext_count.items(), key=lambda kv: (-kv[1], kv[0])):
            print(f"    {ext:<12} {n}")
    print("=" * 60)
    return EXIT_HITS_FOUND if total > 0 else EXIT_NO_HITS


def _run_scan(ctx: ScanContext, file_list: List[str],
              parallel_files: List[str], serial_files: List[str],
              parallel_workers: int) -> ScanResult:
    """並列 + 直列の走査ループを実行する。"""
    log = ctx.log
    args = ctx.args
    cfg = ctx.cfg

    timeout_sec = float(cfg["runtime"].get("per_file_timeout_sec", 0) or 0)
    max_files = 1 if args.first_hit_only else (args.max_files or 0)
    stop_event = threading.Event()

    result = ScanResult(
        total=len(file_list),
        parallel_workers=parallel_workers,
        parallel_count=len(parallel_files),
        serial_count=len(serial_files),
        stop_max=max_files,
    )

    log.info(
        "走査開始: %d ファイル (並列対象=%d / COM 直列=%d, 並列度=%d, タイムアウト=%s) / paths=%s",
        result.total, result.parallel_count, result.serial_count, parallel_workers,
        f"{timeout_sec:g}s" if timeout_sec > 0 else "無効",
        ctx.paths,
    )

    try:
        from tqdm import tqdm  # type: ignore
    except ImportError:
        tqdm = lambda *a, **kw: _NoopBar(*a, **kw)  # type: ignore

    pbar = tqdm(total=result.total, desc="scan", unit="file", disable=args.quiet)
    console_lock = threading.Lock()
    start = time.monotonic()

    def _on_result(res: Tuple[str, Optional[List], Optional[str], Optional[str]]) -> None:
        path, hits, skip_reason, error = res
        if error:
            result.error_results.append(FileResult(path=path, error=error))
            log.debug("error: %s: %s", path, error)
        elif skip_reason:
            result.skipped += 1
            result.skip_files.setdefault(skip_reason, []).append(path)
        elif hits:
            fr = FileResult(path=path, hits=hits)
            result.hit_results.append(fr)
            result.hits_total += len(hits)
            if cfg["output"]["console"]:
                from reporter.console import render_console
                with console_lock:
                    render_console(fr, quiet=args.quiet)
        pbar.update(1)
        if max_files and len(result.hit_results) >= max_files and not stop_event.is_set():
            stop_event.set()
            log.info("ヒットファイル %d 件に達したため走査を打ち切ります。", max_files)

    def _process_with_timeout(path: str, reg: ExtractorRegistry) -> Tuple:
        if stop_event.is_set():
            return path, None, "stopped", None
        if timeout_sec <= 0:
            return _process_file(path, reg, ctx.searcher, cache=ctx.cache)
        with cf.ThreadPoolExecutor(max_workers=1,
                                   thread_name_prefix="docgrep-to") as ex:
            fut = ex.submit(_process_file, path, reg, ctx.searcher, cache=ctx.cache)
            try:
                return fut.result(timeout=timeout_sec)
            except cf.TimeoutError:
                ex.shutdown(wait=False, cancel_futures=True)
                if ctx.com is not None and reg is ctx.serial_registry:
                    try:
                        ctx.com.recover_all()
                        log.warning("タイムアウト後に COM インスタンスを再生成しました")
                    except Exception:
                        pass
                return path, None, None, f"timeout_error: exceeded {timeout_sec:g}s"

    try:
        if parallel_files:
            if parallel_workers > 1:
                with cf.ThreadPoolExecutor(max_workers=parallel_workers,
                                           thread_name_prefix="docgrep") as ex:
                    futures = [ex.submit(_process_with_timeout, p, ctx.parallel_registry)
                               for p in parallel_files]
                    try:
                        if args.ordered_output:
                            for fut in futures:
                                if stop_event.is_set():
                                    fut.cancel()
                                    continue
                                _on_result(fut.result())
                        else:
                            for fut in cf.as_completed(futures):
                                if stop_event.is_set():
                                    for f in futures:
                                        if not f.done():
                                            f.cancel()
                                    break
                                _on_result(fut.result())
                    except KeyboardInterrupt:
                        for f in futures:
                            f.cancel()
                        raise
            else:
                for p in parallel_files:
                    if stop_event.is_set():
                        break
                    _on_result(_process_with_timeout(p, ctx.parallel_registry))

        for p in serial_files:
            if stop_event.is_set():
                break
            _on_result(_process_with_timeout(p, ctx.serial_registry))
    except KeyboardInterrupt:
        result.interrupted = True
        log.warning("中断されました。これまでの結果を出力します。")
    finally:
        pbar.close()

    result.elapsed = time.monotonic() - start
    result.stopped = stop_event.is_set() and not result.interrupted
    return result


def _build_summary(cfg: Dict[str, Any], keywords: List[str],
                   result: ScanResult) -> Dict[str, Any]:
    summary: Dict[str, Any] = {
        "走査ファイル数": result.total,
        "ヒットファイル数": len(result.hit_results),
        "ヒット件数合計": result.hits_total,
        "スキップ数": result.skipped,
        "エラー数": len(result.error_results),
        "処理時間(秒)": f"{result.elapsed:.2f}",
        "検索モード": cfg["search"]["mode"],
        "演算子": cfg["search"]["operator"],
        "キーワード": ", ".join(keywords),
    }
    if result.skip_files:
        summary["スキップ内訳"] = ", ".join(
            f"{k}={len(v)}" for k, v in result.skip_files.items()
        )
    if result.interrupted:
        summary["中断"] = "Ctrl+C により途中終了"
    if result.stopped:
        summary["打ち切り"] = f"max_files={result.stop_max} に到達"
    return summary


def _emit_reports(ctx: ScanContext, result: ScanResult,
                  summary: Dict[str, Any], slug: str) -> None:
    cfg = ctx.cfg
    log = ctx.log
    if cfg["output"]["excel"]["enabled"]:
        try:
            from reporter.excel import write_excel
            excel_path = apply_timestamp(cfg["output"]["excel"]["path"], slug)
            ensure_dir(Path(excel_path).resolve().parent)
            write_excel(excel_path, result.hit_results, summary,
                        errors=result.error_results, skipped=result.skip_files)
            log.info("Excel 出力: %s", excel_path)
        except Exception as e:
            log.warning("Excel 出力に失敗: %s", e)
            log.debug("Excel 出力 traceback", exc_info=True)

    html_out_dir: Optional[Path] = None
    latest_html_name: Optional[str] = None

    if cfg["output"]["html"]["enabled"]:
        try:
            from reporter.html import write_html
            html_path = apply_timestamp(cfg["output"]["html"]["path"], slug)
            ensure_dir(Path(html_path).resolve().parent)
            write_html(html_path, result.hit_results, summary, ctx.searcher,
                       errors=result.error_results, skipped=result.skip_files)
            log.info("HTML 出力: %s", html_path)
            html_out_dir = Path(html_path).resolve().parent

            latest_template = cfg["output"]["html"].get("latest_path")
            if latest_template:
                latest_path = apply_timestamp(latest_template, slug)
                if Path(latest_path).resolve() != Path(html_path).resolve():
                    import shutil
                    ensure_dir(Path(latest_path).resolve().parent)
                    shutil.copyfile(html_path, latest_path)
                    log.info("HTML 最新: %s", latest_path)
                # latest が同じディレクトリにあれば一覧の「最新版」セクションに使う
                if Path(latest_path).resolve().parent == html_out_dir:
                    latest_html_name = Path(latest_path).name
        except Exception as e:
            log.warning("HTML 出力に失敗: %s", e)
            log.debug("HTML 出力 traceback", exc_info=True)

    # === レポート一覧 _index.html を更新 ===
    if html_out_dir is not None:
        try:
            from reporter.index import build_report_index
            idx = build_report_index(html_out_dir, latest_html_name=latest_html_name)
            log.info("レポート一覧: %s", idx)
        except Exception as e:
            log.warning("レポート一覧の生成に失敗: %s", e)
            log.debug("Index 生成 traceback", exc_info=True)


def _print_summary(summary: Dict[str, Any]) -> None:
    print()
    print("=" * 60)
    print(" 検索結果サマリ")
    print("=" * 60)
    for k, v in summary.items():
        print(f"  {k:<14}: {v}")
    print("=" * 60)


# =============================================================================
# キャッシュ管理サブコマンド (--cache-stats / --cache-vacuum / --cache-clear)
# =============================================================================

def _handle_cache_command(cfg: Dict[str, Any], args: argparse.Namespace,
                          log: logging.Logger) -> int:
    """--cache-stats / --cache-vacuum / --cache-clear のいずれかを処理して終了。"""
    cache_cfg = cfg["runtime"].get("cache") or {}
    db_path = Path(cache_cfg.get("path", "reports/.docgrep_cache.sqlite"))
    if not db_path.is_file():
        log.error("キャッシュ DB が存在しません: %s", db_path)
        return EXIT_CONFIG_ERROR
    try:
        from cache import SegmentCache
    except Exception as e:
        log.error("キャッシュモジュールの読込に失敗: %s", e)
        return EXIT_CONFIG_ERROR

    if args.cache_clear:
        try:
            db_path.unlink()
            log.info("キャッシュ DB を削除しました: %s", db_path)
        except Exception as e:
            log.error("削除失敗: %s", e)
            return EXIT_CONFIG_ERROR
        return EXIT_HITS_FOUND

    with SegmentCache(db_path) as c:
        if args.cache_vacuum:
            try:
                c.vacuum()
                size_mb = db_path.stat().st_size / (1024 * 1024)
                log.info("VACUUM 完了: DB サイズ %.2f MB", size_mb)
            except Exception as e:
                log.error("VACUUM 失敗: %s", e)
                return EXIT_CONFIG_ERROR
            return EXIT_HITS_FOUND

        # --cache-stats
        cur = c._conn.execute("SELECT COUNT(*) FROM files")
        n_files = cur.fetchone()[0]
        cur = c._conn.execute("SELECT COUNT(*) FROM segments")
        n_segs = cur.fetchone()[0]
        size_mb = db_path.stat().st_size / (1024 * 1024)
        print()
        print("=" * 60)
        print(" キャッシュ統計")
        print("=" * 60)
        print(f"  DB パス         : {db_path}")
        print(f"  DB サイズ       : {size_mb:.2f} MB")
        print(f"  キャッシュファイル数: {n_files}")
        print(f"  Segment 総数    : {n_segs}")
        if n_files > 0:
            avg = n_segs / n_files
            print(f"  平均 Segment/ファイル: {avg:.1f}")
        print("=" * 60)
    return EXIT_HITS_FOUND


# =============================================================================
# main
# =============================================================================

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

    # --- キャッシュ管理コマンド（早期 return） ---
    if args.cache_stats or args.cache_vacuum or args.cache_clear:
        return _handle_cache_command(cfg, args, log)

    # === ランタイム初期化 ===
    rt = _init_runtime(cfg, args, log)
    if rt is None:
        return EXIT_CONFIG_ERROR
    searcher, keywords = rt

    com, parallel_reg, serial_reg, cache = _init_extractors(cfg, log)

    paths, file_list = _collect_files(cfg)
    if not file_list:
        log.info("走査対象のファイルが見つかりませんでした。paths と extensions を確認してください。")
        if com is not None:
            com.shutdown()
        if cache is not None:
            cache.close()
        return EXIT_NO_HITS

    parallel_files, serial_files = _partition_files(file_list)
    workers = _parallel_workers(cfg)

    # --- DRY RUN（早期 return） ---
    if args.dry_run:
        if com is not None:
            com.shutdown()
        if cache is not None:
            cache.close()
        return _dry_run(file_list, parallel_files, serial_files, workers, paths)

    # === 走査 ===
    ctx = ScanContext(
        cfg=cfg, args=args, log=log, searcher=searcher, keywords=keywords,
        com=com, parallel_registry=parallel_reg, serial_registry=serial_reg,
        cache=cache, paths=paths,
    )
    try:
        result = _run_scan(ctx, file_list, parallel_files, serial_files, workers)
    finally:
        if com is not None:
            com.shutdown()
        if cache is not None:
            log.info("キャッシュ統計: hits=%d, misses=%d, writes=%d",
                     cache.hits, cache.misses, cache.writes)
            cache.close()

    # === 出力 ===
    summary = _build_summary(cfg, keywords, result)
    slug = timestamp_slug()
    _emit_reports(ctx, result, summary, slug)
    _print_summary(summary)

    if result.interrupted:
        return EXIT_INTERRUPTED
    return EXIT_HITS_FOUND if result.hits_total > 0 else EXIT_NO_HITS
