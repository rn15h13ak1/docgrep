"""SegmentCache のテスト。"""
import os
import time

from cache import SegmentCache
from search import Segment


def test_put_then_get_returns_segments(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("hello", encoding="utf-8")
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c:
        segs = [Segment(text="hello", locator="行 1"),
                Segment(text="world", locator="行 2")]
        c.put(str(f), segs)
        out = c.get(str(f))
        assert out is not None
        assert [s.text for s in out] == ["hello", "world"]
        assert [s.locator for s in out] == ["行 1", "行 2"]
        assert c.hits == 1 and c.writes == 1


def test_get_miss_when_file_modified(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("v1", encoding="utf-8")
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c:
        c.put(str(f), [Segment(text="v1", locator="行 1")])
        # 内容を変えて mtime/size を変える
        time.sleep(0.01)
        f.write_text("v2 longer content", encoding="utf-8")
        os.utime(f, None)
        out = c.get(str(f))
        assert out is None  # 変更検知でミス
        assert c.misses >= 1


def test_get_unknown_file_returns_none(tmp_path):
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c:
        assert c.get(str(tmp_path / "nonexistent.txt")) is None


def test_forget_removes_entry(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("hi", encoding="utf-8")
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c:
        c.put(str(f), [Segment(text="hi", locator="行 1")])
        assert c.get(str(f)) is not None
        c.forget(str(f))
        # ファイル自体は存在するが DB エントリは消えた
        c.hits = c.misses = 0
        assert c.get(str(f)) is None


def test_overwrite_replaces_segments(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("hi", encoding="utf-8")
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c:
        c.put(str(f), [Segment(text="old", locator="L1")])
        c.put(str(f), [Segment(text="new1", locator="L1"),
                       Segment(text="new2", locator="L2")])
        out = c.get(str(f))
        assert out is not None
        assert [s.text for s in out] == ["new1", "new2"]


def test_persistence_across_instances(tmp_path):
    f = tmp_path / "doc.txt"
    f.write_text("persist", encoding="utf-8")
    db = tmp_path / "cache.sqlite"
    with SegmentCache(db) as c1:
        c1.put(str(f), [Segment(text="persist", locator="行 1")])
    # 別インスタンスで読み直せる
    with SegmentCache(db) as c2:
        out = c2.get(str(f))
        assert out is not None
        assert out[0].text == "persist"
