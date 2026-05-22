"""抽出ディスパッチ。

拡張子に応じて適切な抽出器を呼び分け、(segments, skip_reason) を返す。

- 既知の Office 形式は専用抽出器を使う
- スコープ外形式 (PDF, .one) はスキップ通知
- 既知のバイナリ／メディア拡張子は即スキップ（"binary" 理由）
- それ以外は **テキストとして抽出を試みる**（中身がバイナリなら "binary" でスキップ）

これにより `.json` / `.xml` / `.html` / 拡張子なし等も、ファイル中身がテキストなら
検索対象に含まれる。
"""
from __future__ import annotations

import os
from typing import List, Optional, Tuple

from search import Segment

XLSX_EXTS = {".xlsx", ".xlsm"}
WORD_EXTS = {".docx", ".doc"}
PPT_EXTS = {".pptx", ".ppt"}
XLS_OLD_EXTS = {".xls"}
SKIP_EXTS = {".pdf", ".one"}  # 初版スコープ外

# 中身を見るまでもなくバイナリと分かる拡張子。テキスト抽出を試みず即スキップ。
BINARY_HINT_EXTS = {
    # 実行形式・ライブラリ
    ".exe", ".dll", ".so", ".dylib", ".bin", ".com", ".sys", ".o", ".obj", ".a", ".lib",
    # アーカイブ
    ".zip", ".tar", ".gz", ".bz2", ".7z", ".rar", ".xz", ".tgz", ".tbz2",
    # 画像
    ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".ico", ".tif", ".tiff", ".psd",
    ".heic", ".heif", ".raw",
    # 音声・動画
    ".mp3", ".mp4", ".avi", ".mov", ".wav", ".flac", ".mkv", ".webm", ".wmv", ".m4a",
    ".m4v", ".aac", ".ogg", ".flv",
    # フォント
    ".ttf", ".otf", ".woff", ".woff2", ".eot",
    # その他
    ".pyc", ".pyo", ".class", ".jar", ".war",
    ".db", ".sqlite", ".sqlite3", ".mdb", ".accdb",
    ".iso", ".dmg", ".vhd", ".vmdk",
}


class ExtractorRegistry:
    def __init__(self, com=None, xlsx_granularity: str = "cell") -> None:
        self.com = com
        self.xlsx_granularity = xlsx_granularity
        from .text import extract_text
        from .xlsx import extract_xlsx
        self._extract_text = extract_text
        self._extract_xlsx = extract_xlsx

    def extract(
        self, path: str
    ) -> Tuple[Optional[List[Segment]], Optional[str]]:
        ext = os.path.splitext(path)[1].lower()

        # スコープ外
        if ext in SKIP_EXTS:
            return None, "scope_out"

        # Office 形式（専用抽出器）
        if ext in XLSX_EXTS:
            return self._extract_xlsx(path, granularity=self.xlsx_granularity), None
        if ext in WORD_EXTS:
            if self.com is None:
                return None, "no_office"
            return self.com.extract_word(path), None
        if ext in PPT_EXTS:
            if self.com is None:
                return None, "no_office"
            return self.com.extract_powerpoint(path), None
        if ext in XLS_OLD_EXTS:
            if self.com is None:
                return None, "no_office"
            return self.com.extract_excel_old(path), None

        # 既知バイナリ拡張子は中身を見ずスキップ
        if ext in BINARY_HINT_EXTS:
            return None, "binary"

        # 残りはテキストとして抽出を試みる（中身がバイナリなら [] が返る）
        segments = self._extract_text(path)
        if not segments:
            return None, "binary"
        return segments, None
