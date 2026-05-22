"""検索エンジン: keyword / regex / fuzzy。

抽出器が返した `List[Segment]` を入力とし、各セグメント内でマッチを取り、
そのセグメントの `locator`（シート名・セル座標・図形名など人間可読の出所）を
ヒットに付与する。keyword AND はファイル全体（全 Segment 連結）で判定する。
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher
from typing import List, Sequence, Tuple

from normalize import normalize
from snippet import make_snippet


@dataclass(frozen=True)
class Segment:
    """抽出器が返すテキスト断片＋出所識別子。"""
    text: str
    locator: str = ""


@dataclass
class Hit:
    snippet: str
    position: int          # Segment 内のオフセット
    matched: str
    locator: str = ""      # 検知箇所（Segment 由来）
    snippet_html: str = ""


@dataclass
class FileResult:
    path: str
    hits: List[Hit] = field(default_factory=list)
    error: str = ""


class Searcher:
    def __init__(
        self,
        mode: str,
        keywords: Sequence[str],
        operator: str,
        case_sensitive: bool,
        normalize_width: bool,
        fuzzy_threshold: float,
        snippet_chars: int,
        max_hits_per_file: int,
    ) -> None:
        if mode not in ("keyword", "regex", "fuzzy"):
            raise ValueError(f"unknown search mode: {mode}")
        if operator not in ("and", "or"):
            raise ValueError(f"unknown operator: {operator}")
        self.mode = mode
        self.operator = operator
        self.case_sensitive = case_sensitive
        self.normalize_width = normalize_width
        self.fuzzy_threshold = fuzzy_threshold
        self.snippet_chars = snippet_chars
        self.max_hits_per_file = max_hits_per_file
        self.raw_keywords: List[str] = [k for k in keywords if k]
        self.norm_keywords: List[str] = [
            normalize(k, case_sensitive, normalize_width) for k in self.raw_keywords
        ]
        self.regexes: List[re.Pattern[str]] = []
        if mode == "regex":
            flags = 0 if case_sensitive else re.IGNORECASE
            for k in self.raw_keywords:
                self.regexes.append(re.compile(k, flags))

    # ---------- public ----------
    def search(self, segments: Sequence[Segment]) -> List[Hit]:
        if not segments or not self.raw_keywords:
            return []
        if self.mode == "regex":
            return self._search_regex(segments)
        # keyword / fuzzy 共通: 正規化済み (text, locator) を作る
        norm_segments: List[Tuple[str, str]] = [
            (normalize(seg.text, self.case_sensitive, self.normalize_width), seg.locator)
            for seg in segments
        ]
        if self.mode == "keyword":
            return self._search_keyword(norm_segments)
        return self._search_fuzzy(norm_segments)

    # ---------- keyword ----------
    def _search_keyword(self, norm_segments: List[Tuple[str, str]]) -> List[Hit]:
        # AND: ファイル全体で全キーワードが出現するかチェック
        if self.operator == "and":
            joined = "\n".join(t for t, _ in norm_segments)
            if not all(kw in joined for kw in self.norm_keywords if kw):
                return []

        hits: List[Hit] = []
        for norm_text, locator in norm_segments:
            if not norm_text:
                continue
            for raw, norm in zip(self.raw_keywords, self.norm_keywords):
                if not norm:
                    continue
                start = 0
                while True:
                    idx = norm_text.find(norm, start)
                    if idx < 0:
                        break
                    hits.append(Hit(
                        snippet=make_snippet(norm_text, idx, idx + len(norm), self.snippet_chars),
                        position=idx,
                        matched=raw,
                        locator=locator,
                    ))
                    if len(hits) >= self.max_hits_per_file:
                        return hits
                    start = idx + max(1, len(norm))
        return hits

    # ---------- regex ----------
    def _search_regex(self, segments: Sequence[Segment]) -> List[Hit]:
        hits: List[Hit] = []
        for seg in segments:
            if not seg.text:
                continue
            for raw, rgx in zip(self.raw_keywords, self.regexes):
                for m in rgx.finditer(seg.text):
                    hits.append(Hit(
                        snippet=make_snippet(seg.text, m.start(), m.end(), self.snippet_chars),
                        position=m.start(),
                        matched=m.group(0) or raw,
                        locator=seg.locator,
                    ))
                    if len(hits) >= self.max_hits_per_file:
                        return hits
        return hits

    # ---------- fuzzy ----------
    def _search_fuzzy(self, norm_segments: List[Tuple[str, str]]) -> List[Hit]:
        hits: List[Hit] = []
        for norm_text, locator in norm_segments:
            if not norm_text:
                continue
            lines = norm_text.splitlines(keepends=True) or [norm_text]
            offset = 0
            for line in lines:
                stripped = line.strip()
                if stripped:
                    for kw in self.norm_keywords:
                        if not kw:
                            continue
                        pos, ratio = _best_window_match(stripped, kw, self.fuzzy_threshold)
                        if pos >= 0:
                            line_start = line.find(stripped)
                            absolute = offset + line_start + pos
                            hits.append(Hit(
                                snippet=make_snippet(
                                    norm_text,
                                    absolute,
                                    absolute + len(kw),
                                    self.snippet_chars,
                                ),
                                position=absolute,
                                matched=f"{kw} (sim={ratio:.2f})",
                                locator=locator,
                            ))
                            if len(hits) >= self.max_hits_per_file:
                                return hits
                offset += len(line)
        return hits


def _best_window_match(haystack: str, needle: str, threshold: float) -> Tuple[int, float]:
    n = len(needle)
    if n == 0 or not haystack:
        return -1, 0.0
    if len(haystack) <= n * 2:
        ratio = SequenceMatcher(None, needle, haystack).ratio()
        if ratio >= threshold:
            return 0, ratio
        return -1, 0.0
    step = max(1, n // 2)
    best_ratio = 0.0
    best_pos = -1
    window = max(n, int(n * 1.5))
    for i in range(0, len(haystack) - window + 1, step):
        chunk = haystack[i:i + window]
        ratio = SequenceMatcher(None, needle, chunk).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_pos = i
        if best_ratio >= 0.99:
            break
    if best_ratio >= threshold:
        return best_pos, best_ratio
    return -1, 0.0
