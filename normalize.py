"""文字列正規化（大文字小文字・全角半角）。"""
from __future__ import annotations

import unicodedata


def normalize(text: str, case_sensitive: bool, normalize_width: bool) -> str:
    """検索用に正規化した文字列を返す。

    NFKC は半角→全角・全角→半角の差異を吸収するが、合字の分解で長さが
    変わる場合がある。本ツールでは正規化後のテキストに対してマッチ位置
    を取り、同じく正規化後の文字列からスニペットを切るため位置ずれは生じない。
    """
    if normalize_width:
        text = unicodedata.normalize("NFKC", text)
    if not case_sensitive:
        text = text.casefold()
    return text
