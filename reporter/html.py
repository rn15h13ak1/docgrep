"""HTML レポート出力（jinja2）。ヒット箇所をハイライト表示する。"""
from __future__ import annotations

import html
import re
from typing import Dict, Iterable, List, Optional

from jinja2 import Template

from search import FileResult, Searcher
from utils import display_path


_TEMPLATE = Template("""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>docgrep 検索結果</title>
  <style>
    :root { color-scheme: light; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif; margin: 20px; color: #222; background: #fff; }
    h1 { font-size: 20px; margin-bottom: 8px; }
    h2 { font-size: 16px; margin: 24px 0 8px; padding-top: 16px; border-top: 2px solid #ddd; }
    .summary { background: #f5f5f5; padding: 12px 16px; border-radius: 6px; margin: 12px 0 20px; font-size: 14px; }
    .summary div { margin: 2px 0; }
    .file { border-top: 1px solid #ddd; padding: 14px 0; }
    .path { font-weight: bold; word-break: break-all; }
    .path a { color: #0b67c2; text-decoration: none; }
    .meta { color: #777; font-size: 12px; margin: 2px 0 6px; }
    .hit { margin: 4px 0 4px 16px; font-size: 13px; line-height: 1.6; }
    .hit .loc { display: inline-block; background: #e3f2fd; color: #0b67c2; font-family: Menlo, Consolas, monospace; font-size: 11px; padding: 1px 6px; border-radius: 3px; margin-right: 8px; vertical-align: middle; }
    .hit .snippet { font-family: Menlo, Consolas, "Courier New", monospace; white-space: pre-wrap; }
    .hl { background: #ffeb3b; padding: 0 2px; border-radius: 2px; }
    .empty { color: #999; padding: 20px 0; }
    .errors { background: #fdecea; border-left: 4px solid #d93025; padding: 8px 16px; border-radius: 4px; margin-top: 12px; }
    .error-row { font-family: Menlo, Consolas, monospace; font-size: 12px; padding: 4px 0; word-break: break-all; }
    .error-row .msg { color: #b00020; }
  </style>
</head>
<body>
  <h1>docgrep 検索結果</h1>
  <div class="summary">
    {% for k, v in summary.items() %}
      <div><strong>{{ k }}</strong>: {{ v }}</div>
    {% endfor %}
  </div>
  {% if not file_results %}
    <div class="empty">ヒットしたファイルはありません。</div>
  {% endif %}
  {% for fr in file_results %}
    <div class="file">
      <div class="path"><a href="file:///{{ fr.path_url }}">{{ fr.path_display }}</a></div>
      <div class="meta">{{ fr.hits|length }} hits</div>
      {% for hit in fr.hits %}
        <div class="hit">
          {% if hit.locator %}<span class="loc">{{ hit.locator }}</span>{% endif %}
          <span class="snippet">{{ hit.snippet_html | safe }}</span>
        </div>
      {% endfor %}
    </div>
  {% endfor %}
  {% if errors %}
    <h2>エラー一覧 ({{ errors|length }})</h2>
    <div class="errors">
      {% for fr in errors %}
        <div class="error-row">
          <div>{{ fr.path_display }}</div>
          <div class="msg">{{ fr.error }}</div>
        </div>
      {% endfor %}
    </div>
  {% endif %}
</body>
</html>
""")


def _highlight(snippet: str, patterns: List[re.Pattern[str]]) -> str:
    escaped = html.escape(snippet)
    for pat in patterns:
        escaped = pat.sub(lambda m: f'<span class="hl">{m.group(0)}</span>', escaped)
    return escaped


def _build_highlight_patterns(searcher: Searcher) -> List[re.Pattern[str]]:
    flags = 0 if searcher.case_sensitive else re.IGNORECASE
    patterns: List[re.Pattern[str]] = []
    if searcher.mode == "regex":
        for kw in searcher.raw_keywords:
            try:
                patterns.append(re.compile(kw, flags))
            except re.error:
                continue
    else:
        for kw in searcher.raw_keywords:
            if kw:
                patterns.append(re.compile(re.escape(html.escape(kw)), flags))
    return patterns


def _annotate(fr: FileResult) -> FileResult:
    """テンプレートで使う表示用属性を FileResult に付与する。"""
    fr.path_url = display_path(fr.path)  # type: ignore[attr-defined]
    fr.path_display = display_path(fr.path)  # type: ignore[attr-defined]
    return fr


def write_html(
    out_path: str,
    file_results: Iterable[FileResult],
    summary: Dict[str, object],
    searcher: Searcher,
    errors: Optional[Iterable[FileResult]] = None,
) -> None:
    patterns = _build_highlight_patterns(searcher)
    rendered: List[FileResult] = []
    for fr in file_results:
        if not fr.hits:
            continue
        for hit in fr.hits:
            hit.snippet_html = _highlight(hit.snippet, patterns)
        rendered.append(_annotate(fr))

    err_list = [_annotate(fr) for fr in (errors or []) if fr.error]

    html_str = _TEMPLATE.render(file_results=rendered, summary=summary, errors=err_list)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_str)
