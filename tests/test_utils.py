from datetime import datetime

from utils import apply_timestamp, display_path, ensure_dir, timestamp_slug


def test_timestamp_slug_format():
    dt = datetime(2026, 5, 21, 9, 30, 15)
    assert timestamp_slug(dt) == "20260521-093015"


def test_apply_timestamp_replaces_placeholder():
    assert apply_timestamp("out/result_{ts}.xlsx", "X") == "out/result_X.xlsx"


def test_apply_timestamp_no_placeholder_passthrough():
    assert apply_timestamp("out/static.xlsx", "X") == "out/static.xlsx"


def test_apply_timestamp_empty_input():
    assert apply_timestamp("", "X") == ""


def test_display_path_normalises_backslashes():
    assert display_path("C:\\Users\\foo") == "C:/Users/foo"


def test_display_path_passthrough():
    assert display_path("/already/posix") == "/already/posix"


def test_ensure_dir_creates_recursively(tmp_path):
    target = tmp_path / "a" / "b" / "c"
    ensure_dir(target)
    assert target.is_dir()


def test_ensure_dir_idempotent(tmp_path):
    target = tmp_path / "exists"
    target.mkdir()
    ensure_dir(target)  # 例外にならない
    assert target.is_dir()
