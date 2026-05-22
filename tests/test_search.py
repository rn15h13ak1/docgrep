from search import Segment, Searcher


def _segs(*pairs):
    """テスト用ヘルパー。(text, locator) のタプルから Segment リストを作る。

    locator を省略するなら 1 引数（テキストだけ）の文字列を渡してもよい。
    """
    out = []
    for p in pairs:
        if isinstance(p, tuple):
            out.append(Segment(text=p[0], locator=p[1]))
        else:
            out.append(Segment(text=p, locator=""))
    return out


def _hits(searcher, segments):
    return [(h.matched, h.locator, h.snippet) for h in searcher.search(segments)]


# === keyword: AND / OR / NFKC / case ===

def test_keyword_and_requires_all_across_segments():
    s = Searcher("keyword", ["foo", "bar"], "and", False, True, 0.8, 30, 10)
    # 同一 segment に両方が無くてもファイル全体で揃えば AND ヒット
    segs = _segs(("foo", "A"), ("bar", "B"))
    hits = s.search(segs)
    assert {h.matched for h in hits} == {"foo", "bar"}
    assert {h.locator for h in hits} == {"A", "B"}


def test_keyword_and_fails_when_one_keyword_missing():
    s = Searcher("keyword", ["foo", "bar"], "and", False, True, 0.8, 30, 10)
    assert s.search(_segs("foo only", "still no other keyword")) == []


def test_keyword_or_returns_each_hit_with_locator():
    s = Searcher("keyword", ["foo", "bar"], "or", False, True, 0.8, 30, 10)
    hits = s.search(_segs(("foo here", "X"), ("bar there", "Y")))
    locators = sorted(h.locator for h in hits)
    assert locators == ["X", "Y"]


def test_keyword_nfkc_fullwidth():
    s = Searcher("keyword", ["ABC123"], "or", False, True, 0.8, 30, 10)
    assert _hits(s, _segs(("ＡＢＣ１２３ を含む", "Sheet1!B5")))


def test_keyword_case_sensitive():
    s_ci = Searcher("keyword", ["Hello"], "or", False, True, 0.8, 30, 10)
    s_cs = Searcher("keyword", ["Hello"], "or", True, False, 0.8, 30, 10)
    assert _hits(s_ci, _segs("say HELLO!"))
    assert not s_cs.search(_segs("say HELLO!"))


def test_keyword_japanese_and_across_cells():
    s = Searcher("keyword", ["東京", "六本木"], "and", False, True, 0.8, 30, 10)
    segs = _segs(("東京本社", "Sheet1!B2"), ("六本木支店", "Sheet1!B3"))
    hits = s.search(segs)
    assert len(hits) == 2
    assert {h.locator for h in hits} == {"Sheet1!B2", "Sheet1!B3"}


def test_keyword_max_hits_limit():
    s = Searcher("keyword", ["a"], "or", False, True, 0.8, 5, 3)
    hits = s.search(_segs("aaaaaaaaa"))
    assert len(hits) == 3


# === regex ===

def test_regex_match_keeps_locator():
    s = Searcher("regex", [r"[A-Z]{3}\d{3}"], "or", False, True, 0.8, 30, 10)
    segs = _segs(("see code ABC123", "L1"), ("and XYZ999 here", "L2"))
    hits = s.search(segs)
    matched = sorted((h.matched, h.locator) for h in hits)
    assert matched == [("ABC123", "L1"), ("XYZ999", "L2")]


def test_regex_invalid_raises():
    import pytest
    with pytest.raises(Exception):
        Searcher("regex", ["[unterminated"], "or", False, True, 0.8, 30, 10)


# === fuzzy ===

def test_fuzzy_picks_up_typo_with_locator():
    s = Searcher("fuzzy", ["quick brown fox"], "or", False, True, 0.7, 30, 10)
    segs = _segs(("the quik bron fox jumps over the lazy dog", "para 1"))
    hits = s.search(segs)
    assert hits
    assert hits[0].locator == "para 1"


def test_fuzzy_threshold_blocks_unrelated():
    s = Searcher("fuzzy", ["specific term"], "or", False, True, 0.95, 30, 10)
    assert not s.search(_segs("totally different content here"))


# === 共通 ===

def test_empty_segments_no_hit():
    s = Searcher("keyword", ["foo"], "or", False, True, 0.8, 30, 10)
    assert s.search([]) == []


def test_empty_keywords_no_hit():
    s = Searcher("keyword", [], "or", False, True, 0.8, 30, 10)
    assert s.search(_segs("anything")) == []


def test_unknown_mode_raises():
    import pytest
    with pytest.raises(ValueError):
        Searcher("unknown", ["x"], "or", False, True, 0.8, 30, 10)
