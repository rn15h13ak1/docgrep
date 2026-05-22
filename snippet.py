"""ヒット箇所のスニペット切り出し。"""
from __future__ import annotations


def make_snippet(text: str, start: int, end: int, snippet_chars: int) -> str:
    """text[start:end] の前後 snippet_chars 文字を切り出し、改行を空白に変換して返す。"""
    s = max(0, start - snippet_chars)
    e = min(len(text), end + snippet_chars)
    prefix = "..." if s > 0 else ""
    suffix = "..." if e < len(text) else ""
    chunk = text[s:e].replace("\r", " ").replace("\n", " ").replace("\t", " ")
    return prefix + chunk + suffix
