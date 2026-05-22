"""過去レポート一覧画面 (_index.html) の生成。

reports/ ディレクトリ配下の `search_result_*.html` / `.xlsx` を mtime 降順で並べ、
クリックで対応する HTML / Excel を開ける一覧ページを作る。`latest_path`
(タイムスタンプなし最新版) は別枠で目立たせる。
"""
from __future__ import annotations

import html
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from jinja2 import Template


_INDEX_NAME = "_index.html"


_TEMPLATE = Template("""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <title>docgrep レポート一覧</title>
  <style>
    :root { color-scheme: light; }
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Hiragino Kaku Gothic ProN", "Yu Gothic", sans-serif; margin: 20px; color: #222; background: #fff; }
    h1 { font-size: 20px; margin-bottom: 4px; }
    .updated { color: #777; font-size: 12px; margin-bottom: 20px; }
    h2 { font-size: 15px; margin: 24px 0 8px; padding-top: 16px; border-top: 1px solid #ddd; }
    .latest-box { background: #e3f2fd; border-left: 4px solid #0b67c2; padding: 10px 16px; border-radius: 4px; margin: 12px 0; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th, td { padding: 6px 8px; text-align: left; border-bottom: 1px solid #eee; }
    th { background: #fafbfc; color: #555; font-weight: 600; font-size: 12px; }
    tr:hover td { background: #fbfcfd; }
    a { color: #0b67c2; text-decoration: none; }
    a:hover { text-decoration: underline; }
    .empty { color: #999; padding: 16px 0; }
    .meta { color: #888; font-size: 12px; font-family: Menlo, Consolas, monospace; }
    .sz { color: #888; font-size: 12px; text-align: right; min-width: 70px; }
    .ts { color: #555; font-size: 12px; font-family: Menlo, Consolas, monospace; }
  </style>
</head>
<body>
  <h1>docgrep レポート一覧</h1>
  <div class="updated">最終更新: {{ generated_at }} / 対象ディレクトリ: {{ dir_display }}</div>

  {% if latest %}
    <div class="latest-box">
      <strong>最新版 (固定パス)</strong>:
      <a href="{{ latest.name }}">{{ latest.name }}</a>
      <span class="meta">({{ latest.size }})</span>
    </div>
  {% endif %}

  <h2>HTML レポート ({{ html_entries|length }})</h2>
  {% if html_entries %}
  <table>
    <thead>
      <tr><th>ファイル</th><th>タイムスタンプ</th><th class="sz">サイズ</th></tr>
    </thead>
    <tbody>
      {% for e in html_entries %}
        <tr>
          <td><a href="{{ e.name }}">{{ e.name }}</a></td>
          <td class="ts">{{ e.ts }}</td>
          <td class="sz">{{ e.size }}</td>
        </tr>
      {% endfor %}
    </tbody>
  </table>
  {% else %}
    <div class="empty">HTML レポートはまだありません。</div>
  {% endif %}

  <h2>Excel レポート ({{ excel_entries|length }})</h2>
  {% if excel_entries %}
  <table>
    <thead>
      <tr><th>ファイル</th><th>タイムスタンプ</th><th class="sz">サイズ</th></tr>
    </thead>
    <tbody>
      {% for e in excel_entries %}
        <tr>
          <td><a href="{{ e.name }}">{{ e.name }}</a></td>
          <td class="ts">{{ e.ts }}</td>
          <td class="sz">{{ e.size }}</td>
        </tr>
      {% endfor %}
    </tbody>
  </table>
  {% else %}
    <div class="empty">Excel レポートはまだありません。</div>
  {% endif %}
</body>
</html>
""")


def _human_bytes(n: int) -> str:
    if n < 1024:
        return f"{n} B"
    units = ["KB", "MB", "GB"]
    size = float(n)
    for u in units:
        size /= 1024
        if size < 1024:
            return f"{size:.1f} {u}"
    return f"{size:.1f} TB"


def _format_ts(mtime: float) -> str:
    return datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")


def _list_entries(dir_path: Path, suffix: str, exclude_names: List[str]
                  ) -> List[dict]:
    """dir_path 配下から拡張子一致 + 除外名以外を mtime 降順で返す。"""
    if not dir_path.is_dir():
        return []
    items = []
    for p in dir_path.iterdir():
        if not p.is_file():
            continue
        if p.suffix.lower() != suffix.lower():
            continue
        if p.name in exclude_names:
            continue
        try:
            st = p.stat()
        except OSError:
            continue
        items.append({
            "name": p.name,
            "size": _human_bytes(st.st_size),
            "ts": _format_ts(st.st_mtime),
            "mtime": st.st_mtime,
        })
    items.sort(key=lambda e: e["mtime"], reverse=True)
    return items


def build_report_index(dir_path: Path,
                       latest_html_name: Optional[str] = None) -> Path:
    """dir_path 配下に `_index.html` を生成して返す。

    latest_html_name: `latest_path` で出力された固定名（あれば一覧から除外し、
    別枠の「最新版」セクションに表示する）。
    """
    dir_path = Path(dir_path)
    out = dir_path / _INDEX_NAME
    excludes_html = [_INDEX_NAME]
    if latest_html_name:
        excludes_html.append(latest_html_name)

    html_entries = _list_entries(dir_path, ".html", excludes_html)
    excel_entries = _list_entries(dir_path, ".xlsx", [])

    latest = None
    if latest_html_name:
        latest_path = dir_path / latest_html_name
        if latest_path.is_file():
            try:
                st = latest_path.stat()
                latest = {
                    "name": latest_html_name,
                    "size": _human_bytes(st.st_size),
                    "ts": _format_ts(st.st_mtime),
                }
            except OSError:
                latest = None

    rendered = _TEMPLATE.render(
        generated_at=_format_ts(datetime.now().timestamp()),
        dir_display=html.escape(str(dir_path.resolve())),
        latest=latest,
        html_entries=html_entries,
        excel_entries=excel_entries,
    )
    dir_path.mkdir(parents=True, exist_ok=True)
    out.write_text(rendered, encoding="utf-8")
    return out
