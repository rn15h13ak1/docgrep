"""テキストファイル抽出（文字コード自動判定 + バイナリ判別）。

ファイル内容がテキストかどうかを先頭バイトで判別し、テキストなら行単位で
Segment 化する（locator: "行 N"）。空行はスキップ。
拡張子に依存しないため、`.json` / `.xml` / `.html` / 拡張子なしファイル等も
中身がテキストなら検索対象になる。
"""
from __future__ import annotations

from typing import List

from charset_normalizer import from_path

from search import Segment


_HEAD_SAMPLE_BYTES = 8192

# UTF-16 LE / UTF-16 BE / UTF-8 の BOM
_TEXT_BOMS = (b"\xff\xfe", b"\xfe\xff", b"\xef\xbb\xbf")


def looks_like_text(path: str) -> bool:
    """先頭バイトを見てテキストファイルらしいかを高速判定する。

    - BOM が付いていれば確実にテキスト（UTF-16 を弾かないように先にチェック）
    - NUL バイトを含めば一律バイナリ扱い（一般的なバイナリの目印）
    - 空ファイルはバイナリ扱い（検索対象なし）
    """
    try:
        with open(path, "rb") as f:
            head = f.read(_HEAD_SAMPLE_BYTES)
    except OSError:
        return False
    if not head:
        return False
    if head.startswith(_TEXT_BOMS):
        return True
    if b"\x00" in head:
        return False
    return True


def extract_text(path: str) -> List[Segment]:
    if not looks_like_text(path):
        return []
    try:
        result = from_path(path).best()
    except Exception:
        return []
    if result is None:
        return []
    text = str(result)
    if not text:
        return []
    segments: List[Segment] = []
    for i, line in enumerate(text.splitlines(), start=1):
        if not line.strip():
            continue
        segments.append(Segment(text=line, locator=f"行 {i}"))
    return segments
