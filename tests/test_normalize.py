from normalize import normalize


def test_nfkc_fullwidth_to_halfwidth():
    assert normalize("ＡＢＣ１２３", case_sensitive=False, normalize_width=True) == "abc123"


def test_case_fold():
    assert normalize("HelloWorld", case_sensitive=False, normalize_width=False) == "helloworld"


def test_case_preserved_when_sensitive():
    assert normalize("Hello", case_sensitive=True, normalize_width=True) == "Hello"


def test_no_normalize_width_keeps_fullwidth():
    assert normalize("ＡＢＣ", case_sensitive=True, normalize_width=False) == "ＡＢＣ"
