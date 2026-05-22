"""テキストファイル抽出（文字コード自動判定 + バイナリ判別）。

行単位で Segment を作り、locator に行番号 ("行 N") を入れる。
空行はノイズになるためスキップ（Segment 数の節約にもなる）。
拡張子に依存しないため、`.json` / `.xml` / `.html` / 拡張子なしファイル等も
中身がテキストなら検索対象になる。

性能のため、まず utf-8-sig（BOM 自動処理込みの UTF-8）でデコードを試み、
それでデコードできない場合だけ charset-normalizer に委ねる。実環境のテキストは
大半が UTF-8 / UTF-8 BOM なので、charset-normalizer の重い推定処理を回避できる。
"""
from __future__ import annotations

from typing import List, Optional

from charset_normalizer import from_path

from search import Segment


_HEAD_SAMPLE_BYTES = 8192

# UTF-16 LE / UTF-16 BE / UTF-8 の BOM
_TEXT_BOMS = (b"\xff\xfe", b"\xfe\xff", b"\xef\xbb\xbf")


def looks_like_text(path: str) -> bool:
    """先頭バイトを見てテキストファイルらしいかを高速判定する。

    - BOM が付いていれば確実にテキスト（UTF-16 を弾かないように先にチェック）
    - NUL バイトを含めば一律バイナリ扱い
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


def _read_utf8_fast(path: str) -> Optional[str]:
    """UTF-8 / UTF-8 BOM を高速にデコードする。失敗時は None を返す。

    UTF-16 BOM 付きや Shift-JIS などは UnicodeDecodeError になり None。
    呼び出し側は None なら charset-normalizer にフォールバックする。
    """
    try:
        # utf-8-sig: BOM があれば取り除き、無くてもエラーにならない
        with open(path, "r", encoding="utf-8-sig", errors="strict") as f:
            return f.read()
    except UnicodeDecodeError:
        return None
    except OSError:
        # 呼び出し側で空リスト扱い
        return ""


def extract_text(path: str) -> List[Segment]:
    if not looks_like_text(path):
        return []

    # 1) UTF-8 / UTF-8 BOM ファストパス
    text = _read_utf8_fast(path)

    # 2) ダメなら charset-normalizer にフォールバック
    if text is None:
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
