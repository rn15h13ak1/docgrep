"""text 抽出器のテスト（行番号 locator + バイナリ判定）。"""
from extractors.text import extract_text, looks_like_text


# === 行番号 locator ===

def test_one_segment_per_non_empty_line(tmp_path):
    p = tmp_path / "sample.md"
    p.write_text("# Title\n\nHello world\nSecond line\n\n  \nFourth\n", encoding="utf-8")
    segs = extract_text(str(p))
    locs = [s.locator for s in segs]
    texts = [s.text for s in segs]
    assert locs == ["行 1", "行 3", "行 4", "行 7"]
    assert texts == ["# Title", "Hello world", "Second line", "Fourth"]


def test_crlf_endings_handled(tmp_path):
    p = tmp_path / "sample.txt"
    p.write_bytes(b"line one\r\nline two\r\n")
    segs = extract_text(str(p))
    assert [s.locator for s in segs] == ["行 1", "行 2"]
    assert [s.text for s in segs] == ["line one", "line two"]


def test_empty_file_returns_empty_list(tmp_path):
    p = tmp_path / "empty.txt"
    p.write_text("", encoding="utf-8")
    assert extract_text(str(p)) == []


# === 拡張子非依存（中身判定） ===

def test_extensionless_text_file_is_extracted(tmp_path):
    p = tmp_path / "README"   # 拡張子なし
    p.write_text("hello\nworld\n", encoding="utf-8")
    segs = extract_text(str(p))
    assert [s.text for s in segs] == ["hello", "world"]


def test_json_file_treated_as_text(tmp_path):
    p = tmp_path / "data.json"
    p.write_text('{"key": "value"}\n', encoding="utf-8")
    segs = extract_text(str(p))
    assert any("key" in s.text for s in segs)


# === バイナリ判定 ===

def test_binary_file_with_null_byte_returns_empty(tmp_path):
    p = tmp_path / "binary.bin"
    p.write_bytes(b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x01")
    assert extract_text(str(p)) == []


def test_utf16_with_bom_is_text(tmp_path):
    # UTF-16 LE BOM + ASCII テキスト（NUL を含むが BOM があるのでテキスト扱い）
    p = tmp_path / "utf16.txt"
    content = "Hello\nWorld\n".encode("utf-16-le")
    p.write_bytes(b"\xff\xfe" + content)
    assert looks_like_text(str(p))
    segs = extract_text(str(p))
    texts = [s.text for s in segs]
    assert "Hello" in "".join(texts) or any("Hello" in t for t in texts)


def test_utf8_bom_is_text(tmp_path):
    p = tmp_path / "bom.txt"
    p.write_bytes(b"\xef\xbb\xbfhello world")
    assert looks_like_text(str(p))


def test_utf8_fast_path_strips_bom(tmp_path):
    # utf-8-sig は BOM を除いて読み出すので、Segment の text には BOM が含まれない
    p = tmp_path / "bom.txt"
    p.write_bytes(b"\xef\xbb\xbf\xe6\x9d\xb1\xe4\xba\xac\n\xe5\xa4\xa7\xe9\x98\xaa")
    segs = extract_text(str(p))
    assert [s.text for s in segs] == ["東京", "大阪"]


def test_shift_jis_falls_back_to_charset_normalizer(tmp_path):
    # CP932 (Shift-JIS) は utf-8-sig で UnicodeDecodeError → charset-normalizer 経路へ。
    # サンプルが短すぎると推定が不安定なので、ある程度長めの本文を入れる。
    p = tmp_path / "sjis.txt"
    content = (
        "東京タワーは1958年に完成しました。\n"
        "大阪城は16世紀後半に建てられた歴史的建造物です。\n"
        "京都には多くの寺院や神社が点在しています。\n"
        "横浜は1859年に開港し、貿易の中心地として発展しました。\n"
    )
    p.write_bytes(content.encode("cp932"))
    segs = extract_text(str(p))
    # fallback 経路を通って何らかの行がデコードできていれば OK
    # （charset-normalizer の推定精度はバージョン依存のため、内容そのものは検証しない）
    assert segs, "fallback 経路でデコードできるべき"


def test_looks_like_text_rejects_pure_binary(tmp_path):
    p = tmp_path / "blob.dat"
    p.write_bytes(bytes(range(256)))
    assert not looks_like_text(str(p))


def test_looks_like_text_empty_file_is_false(tmp_path):
    p = tmp_path / "empty.bin"
    p.write_bytes(b"")
    assert not looks_like_text(str(p))
