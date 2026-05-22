"""reporter/index.build_report_index のテスト。"""
import os
import time
from pathlib import Path

from reporter.index import build_report_index


def _touch(path: Path, content: str = "x", mtime: float = None) -> Path:
    path.write_text(content, encoding="utf-8")
    if mtime is not None:
        os.utime(path, (mtime, mtime))
    return path


def test_index_lists_html_and_xlsx_by_mtime_desc(tmp_path):
    base = tmp_path / "reports"
    base.mkdir()
    now = time.time()
    a = _touch(base / "search_result_20260101-000000.html", mtime=now - 200)
    b = _touch(base / "search_result_20260102-000000.html", mtime=now - 100)
    c = _touch(base / "search_result_20260101-000000.xlsx", mtime=now - 150)
    out = build_report_index(base)

    assert out == base / "_index.html"
    text = out.read_text(encoding="utf-8")
    # 両方の HTML が含まれる
    assert "search_result_20260101-000000.html" in text
    assert "search_result_20260102-000000.html" in text
    assert "search_result_20260101-000000.xlsx" in text
    # 新しいものが先（b → a の順で登場）
    idx_b = text.find("search_result_20260102-000000.html")
    idx_a = text.find("search_result_20260101-000000.html")
    assert idx_b < idx_a, "新しい mtime のファイルが上に来ない"


def test_index_excludes_itself(tmp_path):
    base = tmp_path / "reports"
    base.mkdir()
    _touch(base / "search_result_x.html")
    # 既存の _index.html を置く → 次回生成でも _index.html は一覧に出ない
    _touch(base / "_index.html", content="<html></html>")
    out = build_report_index(base)
    text = out.read_text(encoding="utf-8")
    # _index.html が a タグの href としてリストに含まれない
    # （タイトル文字列としては「docgrep レポート一覧」だけで OK）
    assert 'href="_index.html"' not in text


def test_index_latest_section(tmp_path):
    base = tmp_path / "reports"
    base.mkdir()
    _touch(base / "search_result_20260101-000000.html")
    _touch(base / "search_result_latest.html")
    out = build_report_index(base, latest_html_name="search_result_latest.html")
    text = out.read_text(encoding="utf-8")
    # latest-box セクションが含まれる
    assert "latest-box" in text
    assert "search_result_latest.html" in text
    # 通常一覧からは latest_html_name が除外される（href= で1回だけ＝最新枠分のみ）
    assert text.count('href="search_result_latest.html"') == 1


def test_index_handles_empty_dir(tmp_path):
    base = tmp_path / "reports"
    base.mkdir()
    out = build_report_index(base)
    text = out.read_text(encoding="utf-8")
    assert "HTML レポートはまだありません" in text
    assert "Excel レポートはまだありません" in text


def test_index_skips_non_matching_files(tmp_path):
    base = tmp_path / "reports"
    base.mkdir()
    _touch(base / "search_result.html")
    _touch(base / "notes.txt")        # 拡張子違い → 出ない
    _touch(base / "config.yaml")      # 同上
    out = build_report_index(base)
    text = out.read_text(encoding="utf-8")
    assert "search_result.html" in text
    assert "notes.txt" not in text
    assert "config.yaml" not in text
