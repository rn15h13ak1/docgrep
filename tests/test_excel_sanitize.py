"""reporter/excel._sanitize / _row のテスト（openpyxl 不要）。"""
from reporter.excel import _row, _sanitize


def test_sanitize_strips_control_chars():
    # \x07 (BEL), \x0b (VT), \x1f (US) はすべて空白に
    src = "Hello\x07World\x0bTest\x1fEnd"
    assert _sanitize(src) == "Hello World Test End"


def test_sanitize_keeps_tab_lf_cr():
    # \t \n \r は許可されている
    src = "line1\nline2\tcol\rend"
    assert _sanitize(src) == src


def test_sanitize_passthrough_non_string():
    assert _sanitize(42) == 42
    assert _sanitize(None) is None
    assert _sanitize(3.14) == 3.14


def test_row_sanitizes_each_cell():
    # NUL (\x00) を含む値も openpyxl で不可、サニタイズされる
    cells = _row("a", "b\x00c", 42, None, "d\x07e")
    assert cells == ["a", "b c", 42, None, "d e"]
