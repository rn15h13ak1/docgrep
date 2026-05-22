"""cli.py のヘルパ関数を fake registry で動かす統合テスト。

main() 全体を呼ぶと selfcheck / argparse / 副作用が多いため、Sprint D で
分離された各段階（_collect_files / _partition_files / _run_scan /
_dry_run / _handle_cache_command）を直接ドライブする。
"""
import argparse
from pathlib import Path
from typing import List, Optional, Tuple

import pytest

import cli
from cli import (
    EXIT_HITS_FOUND,
    EXIT_NO_HITS,
    ScanContext,
    _build_summary,
    _collect_files,
    _dry_run,
    _handle_cache_command,
    _partition_files,
    _run_scan,
)
from config import DEFAULT_CONFIG, load_config
from copy import deepcopy
from search import FileResult, Hit, Searcher, Segment


# ------------------------------------------------------------------ fakes ----

class FakeRegistry:
    """ExtractorRegistry のスタブ。path → segments のマップを返すだけ。"""

    def __init__(self, mapping: dict, skip: Optional[dict] = None) -> None:
        self.mapping = mapping
        self.skip = skip or {}

    def extract(self, path: str) -> Tuple[Optional[List[Segment]], Optional[str]]:
        if path in self.skip:
            return None, self.skip[path]
        segs = self.mapping.get(path)
        if segs is None:
            return None, None
        return segs, None


def _make_args(**overrides) -> argparse.Namespace:
    """_run_scan が要求する args 属性を持つ Namespace を作る。"""
    defaults = dict(
        verbose=False, quiet=True, ordered_output=False,
        max_files=0, first_hit_only=False, dry_run=False,
    )
    defaults.update(overrides)
    return argparse.Namespace(**defaults)


def _make_searcher(keywords=("hello",), mode="keyword", operator="or"):
    return Searcher(
        mode=mode, keywords=list(keywords), operator=operator,
        case_sensitive=False, normalize_width=True,
        fuzzy_threshold=0.8, snippet_chars=30, max_hits_per_file=10,
    )


def _make_cfg(**runtime_overrides) -> dict:
    cfg = deepcopy(DEFAULT_CONFIG)
    cfg["runtime"].update(runtime_overrides)
    cfg["output"]["console"] = False
    cfg["output"]["excel"]["enabled"] = False
    cfg["output"]["html"]["enabled"] = False
    return cfg


# --------------------------------------------------------- _collect_files ----

def test_collect_files_picks_up_paths_and_filters(tmp_path):
    (tmp_path / "a.txt").write_text("hello", encoding="utf-8")
    (tmp_path / "b.bin").write_bytes(b"\x00\x01")
    sub = tmp_path / "sub"
    sub.mkdir()
    (sub / "c.txt").write_text("nested", encoding="utf-8")

    cfg = _make_cfg()
    cfg["paths"] = [str(tmp_path)]
    cfg["extensions"] = [".txt"]
    paths, files = _collect_files(cfg)
    names = sorted(Path(p).name for p in files)
    assert names == ["a.txt", "c.txt"]
    assert str(tmp_path) in paths


def test_collect_files_appends_onenote_export_dir(tmp_path):
    onenote = tmp_path / "onenote_export"
    onenote.mkdir()
    (onenote / "x.docx").write_text("dummy", encoding="utf-8")
    cfg = _make_cfg()
    cfg["paths"] = []
    cfg["onenote_export_dir"] = str(onenote)
    cfg["extensions"] = [".docx"]
    paths, files = _collect_files(cfg)
    assert str(onenote) in paths


# ------------------------------------------------------- _partition_files ----

def test_partition_files_routes_com_extensions_to_serial():
    files = ["a.txt", "b.xlsx", "c.docx", "d.doc", "e.ppt", "f.csv"]
    parallel, serial = _partition_files(files)
    assert "a.txt" in parallel and "b.xlsx" in parallel and "f.csv" in parallel
    assert set(serial) == {"c.docx", "d.doc", "e.ppt"}


# ----------------------------------------------------------------- _run_scan -

def test_run_scan_aggregates_hits_skips_errors(tmp_path):
    files = [str(tmp_path / f"f{i}.txt") for i in range(3)]
    for f in files:
        Path(f).write_text("placeholder", encoding="utf-8")

    fake = FakeRegistry(
        mapping={
            files[0]: [Segment(text="hello world", locator="行 1")],
            files[1]: [Segment(text="no match here", locator="行 1")],
            # files[2] は extract が None → スキップ理由付き
        },
        skip={files[2]: "binary"},
    )
    searcher = _make_searcher(("hello",))
    ctx = ScanContext(
        cfg=_make_cfg(), args=_make_args(),
        log=__import__("logging").getLogger("docgrep"),
        searcher=searcher, keywords=["hello"],
        parallel_registry=fake, serial_registry=fake,
        paths=[str(tmp_path)],
    )
    result = _run_scan(ctx, files, parallel_files=files, serial_files=[], parallel_workers=1)
    assert result.total == 3
    assert result.hits_total == 1
    assert len(result.hit_results) == 1
    assert "binary" in result.skip_files
    assert result.skip_files["binary"] == [files[2]]


def test_run_scan_stops_on_max_files(tmp_path):
    files = [str(tmp_path / f"hit{i}.txt") for i in range(5)]
    for f in files:
        Path(f).write_text("hello", encoding="utf-8")

    fake = FakeRegistry(
        mapping={f: [Segment(text="hello", locator="行 1")] for f in files}
    )
    searcher = _make_searcher(("hello",))
    ctx = ScanContext(
        cfg=_make_cfg(), args=_make_args(max_files=2),
        log=__import__("logging").getLogger("docgrep"),
        searcher=searcher, keywords=["hello"],
        parallel_registry=fake, serial_registry=fake,
    )
    result = _run_scan(ctx, files, parallel_files=files, serial_files=[], parallel_workers=1)
    # 2 件で止まる。それ以降の files は処理されない/stopped 扱い
    assert len(result.hit_results) >= 2  # workers=1 なら厳密に 2
    assert result.stopped is True


def test_run_scan_records_extract_error_message(tmp_path):
    f = str(tmp_path / "bad.txt")
    Path(f).write_text("x", encoding="utf-8")

    class _RaisingRegistry:
        def extract(self, path):
            raise RuntimeError("boom")

    ctx = ScanContext(
        cfg=_make_cfg(), args=_make_args(),
        log=__import__("logging").getLogger("docgrep"),
        searcher=_make_searcher(("x",)), keywords=["x"],
        parallel_registry=_RaisingRegistry(), serial_registry=_RaisingRegistry(),
    )
    result = _run_scan(ctx, [f], parallel_files=[f], serial_files=[], parallel_workers=1)
    assert len(result.error_results) == 1
    assert "extract_error" in result.error_results[0].error


# --------------------------------------------------------- _build_summary ---

def test_build_summary_includes_skip_breakdown(tmp_path):
    cfg = _make_cfg()
    from cli import ScanResult
    result = ScanResult(
        total=10, hits_total=3, skipped=4,
        hit_results=[FileResult(path="x")],
        error_results=[FileResult(path="e", error="extract_error: x")],
        skip_files={"binary": ["a", "b"], "scope_out": ["c", "d"]},
        elapsed=1.23,
    )
    summary = _build_summary(cfg, ["foo", "bar"], result)
    assert summary["走査ファイル数"] == 10
    assert summary["ヒット件数合計"] == 3
    assert "binary=2" in summary["スキップ内訳"]
    assert "scope_out=2" in summary["スキップ内訳"]


# ------------------------------------------------------------------ _dry_run

def test_dry_run_returns_no_hits_for_empty(capsys):
    rc = _dry_run(file_list=[], parallel_files=[], serial_files=[],
                  parallel_workers=2, paths=["/nonexistent"])
    assert rc == EXIT_NO_HITS


def test_dry_run_prints_breakdown(capsys):
    files = ["a.txt", "b.md", "c.xlsx", "d.docx"]
    rc = _dry_run(file_list=files, parallel_files=["a.txt", "b.md", "c.xlsx"],
                  serial_files=["d.docx"], parallel_workers=4,
                  paths=["/tmp/x"])
    assert rc == EXIT_HITS_FOUND
    out = capsys.readouterr().out
    assert "DRY RUN" in out
    assert ".txt" in out and ".docx" in out
    assert "並列対象      : 3" in out


# ------------------------------------------------------ _handle_cache_command

def _make_cache_args(stats=False, vacuum=False, clear=False):
    return argparse.Namespace(
        cache_stats=stats, cache_vacuum=vacuum, cache_clear=clear,
    )


def test_handle_cache_command_missing_db_returns_config_error(tmp_path, caplog):
    cfg = _make_cfg()
    cfg["runtime"]["cache"] = {"enabled": True,
                                "path": str(tmp_path / "missing.sqlite")}
    rc = _handle_cache_command(cfg, _make_cache_args(stats=True),
                               __import__("logging").getLogger("docgrep"))
    assert rc != 0


def test_handle_cache_command_stats_and_clear(tmp_path):
    # まずキャッシュを 1 件書く
    from cache import SegmentCache
    db = tmp_path / "c.sqlite"
    f = tmp_path / "doc.txt"
    f.write_text("hello", encoding="utf-8")
    with SegmentCache(db) as c:
        c.put(str(f), [Segment(text="hello", locator="行 1")])
    assert db.is_file()

    cfg = _make_cfg()
    cfg["runtime"]["cache"] = {"enabled": True, "path": str(db)}
    log = __import__("logging").getLogger("docgrep")

    rc = _handle_cache_command(cfg, _make_cache_args(stats=True), log)
    assert rc == EXIT_HITS_FOUND

    rc = _handle_cache_command(cfg, _make_cache_args(clear=True), log)
    assert rc == EXIT_HITS_FOUND
    assert not db.is_file()
