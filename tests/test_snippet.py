from snippet import make_snippet


def test_full_context_no_ellipsis():
    text = "abcde"
    s = make_snippet(text, 2, 3, 5)
    assert s == "abcde"


def test_left_ellipsis():
    text = "0123456789abcdef"
    s = make_snippet(text, 10, 11, 3)
    assert s.startswith("...")
    assert "abc" in s


def test_right_ellipsis():
    text = "0123456789abcdef"
    s = make_snippet(text, 2, 3, 3)
    assert s.endswith("...")


def test_newlines_replaced_with_space():
    text = "line1\nline2\nline3"
    s = make_snippet(text, 7, 8, 10)
    assert "\n" not in s
    assert "line" in s
