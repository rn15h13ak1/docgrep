"""HTML レポート出力（jinja2）。ヒット箇所をハイライト表示する。"""
from __future__ import annotations

import html
import json
import os
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
    /* スキップ詳細リンク（モーダル起動ボタン） */
    .skip-link { background: #eef2f7; border: 1px solid #cfd8e3; color: #2d4a6b; padding: 1px 8px; border-radius: 3px; font-size: 12px; font-family: Menlo, Consolas, monospace; cursor: pointer; margin: 0 4px 2px 0; }
    .skip-link:hover { background: #dde6f0; }
    /* モーダル */
    dialog#skip-dialog { max-width: min(900px, 90vw); width: 90vw; max-height: 80vh; border: 1px solid #ccc; border-radius: 8px; padding: 0; box-shadow: 0 10px 30px rgba(0,0,0,0.2); }
    dialog#skip-dialog::backdrop { background: rgba(0,0,0,0.4); }
    .skip-dialog-header { display: flex; justify-content: space-between; align-items: center; padding: 12px 20px; border-bottom: 1px solid #eee; }
    .skip-dialog-header h3 { margin: 0; font-size: 15px; }
    .skip-dialog-list { font-family: Menlo, Consolas, monospace; font-size: 12px; overflow: auto; max-height: 60vh; padding: 12px 20px; }
    .skip-dialog-list .row { padding: 3px 0; word-break: break-all; border-bottom: 1px dotted #eee; }
    .skip-dialog-list .row:last-child { border-bottom: 0; }
    .skip-close { padding: 5px 14px; cursor: pointer; border: 1px solid #cfd8e3; background: #fff; border-radius: 4px; font-size: 12px; }
    .skip-close:hover { background: #f3f5f8; }
    /* フィルタバー */
    .filter-bar { background: #fafbfc; border: 1px solid #e6e8eb; border-radius: 6px; padding: 10px 14px; margin: 12px 0 16px; font-size: 13px; display: flex; flex-wrap: wrap; align-items: center; gap: 10px; }
    .filter-bar input[type="search"] { flex: 1 1 200px; min-width: 200px; padding: 4px 8px; border: 1px solid #cfd8e3; border-radius: 4px; font-size: 13px; }
    .filter-group { display: inline-flex; gap: 6px; align-items: center; flex-wrap: wrap; }
    .filter-chip { padding: 2px 10px; border: 1px solid #cfd8e3; border-radius: 12px; background: #fff; cursor: pointer; font-size: 12px; user-select: none; }
    .filter-chip.active { background: #0b67c2; color: #fff; border-color: #0b67c2; }
    .filter-chip:hover:not(.active) { background: #eef2f7; }
    .filter-count { color: #777; font-size: 12px; margin-left: auto; }
    .file.hidden, .hit.hidden { display: none; }
    .file.dim { opacity: 0.35; }
  </style>
</head>
<body>
  <h1>docgrep 検索結果</h1>
  <div class="summary">
    {% for k, v in summary.items() %}
      <div><strong>{{ k }}</strong>: {{ v }}</div>
    {% endfor %}
    {% if skipped %}
      <div style="margin-top:6px;">
        <strong>スキップ詳細</strong>:
        {% for reason, paths in skipped.items() %}
          <button class="skip-link" data-reason="{{ reason }}">{{ reason }} ({{ paths|length }})</button>
        {% endfor %}
      </div>
    {% endif %}
  </div>

  {% if file_results %}
    <div class="filter-bar" id="filter-bar">
      <input type="search" id="filter-text" placeholder="パス・スニペットで絞り込み（部分一致）">
      <span class="filter-group" id="filter-ext-group">
        <span style="color:#666;">拡張子:</span>
        <span class="filter-chip active" data-ext="">全て</span>
        {% for ext in ext_set %}
          <span class="filter-chip" data-ext="{{ ext }}">{{ ext or '(なし)' }}</span>
        {% endfor %}
      </span>
      <span class="filter-count" id="filter-count"></span>
    </div>
  {% else %}
    <div class="empty">ヒットしたファイルはありません。</div>
  {% endif %}
  {% for fr in file_results %}
    <div class="file" data-ext="{{ fr.ext }}" data-path="{{ fr.path_display | lower }}">
      <div class="path"><a href="file:///{{ fr.path_url }}">{{ fr.path_display }}</a></div>
      <div class="meta">{{ fr.hits|length }} hits</div>
      {% for hit in fr.hits %}
        <div class="hit" data-snippet="{{ hit.snippet_text | lower }}">
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

  {% if skipped %}
    <dialog id="skip-dialog">
      <div class="skip-dialog-header">
        <h3 id="skip-dialog-title">スキップ詳細</h3>
        <button class="skip-close" onclick="document.getElementById('skip-dialog').close()">閉じる</button>
      </div>
      <div id="skip-dialog-list" class="skip-dialog-list"></div>
    </dialog>
    <script>
      const skipData = {{ skipped_json | safe }};
      function escHtml(s) {
        return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
          .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
      }
      document.querySelectorAll('.skip-link').forEach(btn => {
        btn.addEventListener('click', () => {
          const reason = btn.dataset.reason;
          const paths = skipData[reason] || [];
          document.getElementById('skip-dialog-title').textContent =
            'スキップ: ' + reason + ' (' + paths.length + ' 件)';
          document.getElementById('skip-dialog-list').innerHTML =
            paths.map(p => '<div class="row">' + escHtml(p) + '</div>').join('');
          const dlg = document.getElementById('skip-dialog');
          if (typeof dlg.showModal === 'function') {
            dlg.showModal();
          } else {
            // 古いブラウザ向けフォールバック（社内では稀のはず）
            dlg.setAttribute('open', '');
          }
        });
      });
    </script>
  {% endif %}

  {% if file_results %}
  <script>
    (function() {
      const files = Array.from(document.querySelectorAll('.file'));
      const textInput = document.getElementById('filter-text');
      const extChips = document.querySelectorAll('#filter-ext-group .filter-chip');
      const countEl = document.getElementById('filter-count');

      let activeExt = '';
      let activeText = '';

      function apply() {
        let shown = 0;
        const q = activeText.trim().toLowerCase();
        files.forEach(f => {
          const ext = f.dataset.ext || '';
          const path = f.dataset.path || '';
          const extOk = !activeExt || ext === activeExt;
          let textOk = true;
          if (q) {
            // パス OR いずれかの hit の snippet にマッチすれば OK
            textOk = path.includes(q);
            if (!textOk) {
              const hits = f.querySelectorAll('.hit');
              for (const h of hits) {
                if ((h.dataset.snippet || '').includes(q)) { textOk = true; break; }
              }
            }
          }
          if (extOk && textOk) {
            f.classList.remove('hidden');
            shown++;
          } else {
            f.classList.add('hidden');
          }
        });
        countEl.textContent = shown + ' / ' + files.length + ' ファイル表示';
      }

      extChips.forEach(chip => {
        chip.addEventListener('click', () => {
          extChips.forEach(c => c.classList.remove('active'));
          chip.classList.add('active');
          activeExt = chip.dataset.ext || '';
          apply();
        });
      });
      textInput.addEventListener('input', e => {
        activeText = e.target.value || '';
        apply();
      });

      apply();
    })();
  </script>
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
    fr.ext = os.path.splitext(fr.path)[1].lower()  # type: ignore[attr-defined]
    return fr


def write_html(
    out_path: str,
    file_results: Iterable[FileResult],
    summary: Dict[str, object],
    searcher: Searcher,
    errors: Optional[Iterable[FileResult]] = None,
    skipped: Optional[Dict[str, List[str]]] = None,
) -> None:
    patterns = _build_highlight_patterns(searcher)
    rendered: List[FileResult] = []
    ext_set: List[str] = []
    seen_ext: set = set()
    for fr in file_results:
        if not fr.hits:
            continue
        for hit in fr.hits:
            hit.snippet_html = _highlight(hit.snippet, patterns)
            # フィルタの部分一致比較用に小文字スニペットを別属性で持たせる
            hit.snippet_text = hit.snippet  # type: ignore[attr-defined]
        ann = _annotate(fr)
        rendered.append(ann)
        if ann.ext not in seen_ext:  # type: ignore[attr-defined]
            seen_ext.add(ann.ext)  # type: ignore[attr-defined]
            ext_set.append(ann.ext)  # type: ignore[attr-defined]
    ext_set.sort()

    err_list = [_annotate(fr) for fr in (errors or []) if fr.error]
    # display_path で / 統一しつつ JS 側に渡す
    skipped_display: Dict[str, List[str]] = {}
    if skipped:
        for reason, paths in skipped.items():
            skipped_display[reason] = [display_path(p) for p in paths]
    skipped_json = json.dumps(skipped_display, ensure_ascii=False)

    html_str = _TEMPLATE.render(
        file_results=rendered,
        summary=summary,
        errors=err_list,
        skipped=skipped_display,
        skipped_json=skipped_json,
        ext_set=ext_set,
    )
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html_str)
