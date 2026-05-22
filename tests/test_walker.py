from pathlib import Path

from walker import iter_files


def _setup_tree(root: Path):
    (root / "a.txt").write_text("hello")
    (root / "b.log").write_text("log line")
    (root / "c.bin").write_bytes(b"\x00\x01\x02")
    (root / "~$tmp.docx").write_text("excel/word lockfile")
    sub = root / "sub"
    sub.mkdir()
    (sub / "d.txt").write_text("nested")
    excluded = root / ".git"
    excluded.mkdir()
    (excluded / "ignored.txt").write_text("must not appear")


def test_extension_filter(tmp_path):
    _setup_tree(tmp_path)
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[".git"],
        exclude_patterns=["~$*"],
        extensions=[".txt"],
        max_size_mb=10,
    ))
    names = sorted(Path(p).name for p in files)
    assert names == ["a.txt", "d.txt"]


def test_exclude_dirs_skipped(tmp_path):
    _setup_tree(tmp_path)
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[".git"],
        exclude_patterns=[],
        extensions=[".txt", ".log"],
        max_size_mb=10,
    ))
    assert not any(".git" in p for p in files)


def test_exclude_patterns_match_filename(tmp_path):
    _setup_tree(tmp_path)
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[],
        exclude_patterns=["~$*"],
        extensions=[".docx"],
        max_size_mb=10,
    ))
    assert files == []


def test_max_size_filter(tmp_path):
    big = tmp_path / "big.txt"
    big.write_bytes(b"x" * (2 * 1024 * 1024 + 1))  # > 2 MB
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[],
        exclude_patterns=[],
        extensions=[".txt"],
        max_size_mb=1,  # 1 MB
    ))
    assert files == []


def test_single_file_path(tmp_path):
    p = tmp_path / "single.txt"
    p.write_text("data")
    files = list(iter_files(
        [str(p)],
        exclude_dirs=[],
        exclude_patterns=[],
        extensions=[".txt"],
        max_size_mb=10,
    ))
    assert files == [str(p)]


def test_wildcard_disables_extension_filter(tmp_path):
    _setup_tree(tmp_path)
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[".git"],
        exclude_patterns=[],
        extensions=["*"],  # 拡張子無条件
        max_size_mb=10,
    ))
    names = sorted(p.split("/")[-1] for p in files)
    # .git は除外、~$tmp.docx はパターン除外無いのでヒット、.bin/.log/.txt 等を含む
    assert "a.txt" in names
    assert "b.log" in names
    assert "c.bin" in names
    assert "d.txt" in names


def test_empty_extensions_treats_as_wildcard(tmp_path):
    _setup_tree(tmp_path)
    files = list(iter_files(
        [str(tmp_path)],
        exclude_dirs=[".git"],
        exclude_patterns=[],
        extensions=[],
        max_size_mb=10,
    ))
    assert any(p.endswith(".bin") for p in files)


def test_nonexistent_path_skipped(tmp_path):
    files = list(iter_files(
        [str(tmp_path / "does_not_exist")],
        exclude_dirs=[],
        exclude_patterns=[],
        extensions=[".txt"],
        max_size_mb=10,
    ))
    assert files == []
