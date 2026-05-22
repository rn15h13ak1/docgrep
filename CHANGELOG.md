# Changelog

[Keep a Changelog](https://keepachangelog.com/ja/1.1.0/) 形式に概ね準拠。
バージョン番号は付けず、Sprint 単位で区切っています。

## [Unreleased]

## Sprint G — レポート索引

### Added
- `reports/_index.html` を自動生成し、過去の HTML レポートを一覧化（最新順）

## Sprint F — 配布・CI・統合テスト

### Added
- `.github/workflows/test.yml` で push / PR ごとに `pytest` を実行する CI
- `CHANGELOG.md` を新規作成（本ファイル）
- `tests/test_cli_integration.py` — `_run_scan` / `_collect_files` / `_partition_files`
  / `_dry_run` / `_handle_cache_command` を fake registry でドライブ
- `tests/test_menu.py` — `subprocess.run` をモックして `_run_docgrep` / `_run_export` /
  履歴の load/save を検証

## Sprint E — 運用機能

### Added
- `--dry-run`: 走査対象のファイル数・拡張子別件数・並列割当だけ表示して終了
- `--cache-stats`: キャッシュ DB のサイズ / エントリ数 / 平均 Segment 数を表示
- `--cache-vacuum`: 孤児セグメント掃除 + SQLite VACUUM
- `--cache-clear`: キャッシュ DB を全削除
- `menu.py` に入力履歴 (`~/.docgrep_history.json`)。前回キーワード / モード / パスを
  次回起動時の既定値として提示

### Changed
- `argparse` の `--help` を `add_argument_group()` で 5 グループに整理
  （設定・対象パス / 検索条件 / 出力 / 動作制御 / 抽出キャッシュ）

## Sprint D — リファクタとテスト基盤

### Changed
- `cli.main()` を 7 つのヘルパ関数に分割（`_init_runtime` / `_init_extractors` /
  `_collect_files` / `_partition_files` / `_run_scan` / `_build_summary` /
  `_emit_reports`）+ `ScanContext` / `ScanResult` データクラス導入。動作は不変
- `extractors/office_com.py` を `_AppHandle` 基底 + `_WordHandle` / `_ExcelHandle`
  / `_PptHandle` に整理し、3 アプリ分の DispatchEx / Quit / recycle カウンタの
  重複を解消

### Added
- `tests/test_extractors_xlsx.py` — 実 xlsx を openpyxl で生成して
  シート名・セル座標 (`Sheet1!B5`) ・行粒度・コメント (著者付き) ・破損ファイル耐性
  を検証する 8 件のテスト

## Sprint C — 堅牢化と UX

### Added
- `runtime.cache.enabled` / `path` 設定で SQLite 抽出キャッシュを有効化
  （`path + mtime + size` をキーに Segment を永続化、2 回目以降の検索を高速化）
- `--cache` / `--no-cache` / `--cache-path` CLI フラグ
- `xlsx_granularity: cell | row` 設定。`row` は 1 行 = 1 Segment（locator は
  `Sheet1!Row 5`、列番地は失われる代わりに大規模 xlsx でメモリ・時間を節約）
- HTML レポートにクライアントサイド絞り込み UI（パス/スニペット部分一致 +
  拡張子チップ + ライブカウンタ）
- `OfficeCom.recover_all()`: タイムアウト超過時に全インスタンスを破棄して
  再生成可能な状態に戻す
- `config.py` の `_validate_types()`: 型 / 列挙 / 範囲を一括検証して
  `ConfigError` でまとめて報告

## Sprint B — 性能とテスト

### Added
- SQLite 抽出キャッシュの基盤（`cache.py`、`SegmentCache` クラス）。WAL モード、
  hits / misses / writes の統計
- `--max-files N` / `--first-hit-only`: ヒット件数で走査を打ち切る
- `--ordered-output`: 並列処理時もコンソール出力をファイル入力順に保つ
- `tests/test_cache.py` (6 件): キャッシュの put/get、mtime 変更検知、forget、
  上書き、永続化
- `tests/test_reporters.py` (4 件): Excel/HTML の round-trip、制御文字
  サニタイズ、モーダル表示有無

## Sprint A — UX と即効性

### Added
- `runtime.per_file_timeout_sec` を実装。並列ワーカー / 直列 COM フェーズ両方で
  `future.result(timeout=...)` によりタイムアウト検出
- `menu.py` を大幅拡張: 設定ファイル選択、AND/OR、fuzzy しきい値、大小区別 /
  NFKC / verbose の対話指定、検索完了後に HTML レポートを自動オープン

### Fixed
- README の設定ファイル例に含まれていた既定 path のズレを修正
  (`search_result_{ts}.xlsx` → `reports/search_result_{ts}.xlsx`)

## 初期実装 〜 Sprint 番号付け前

### Added
- フェーズ 1〜3 を完了: ファイル走査、Office (Word/PPT/Excel) + テキスト抽出、
  キーワード / 正規表現 / あいまい (difflib) 検索、コンソール / Excel / HTML 出力
- OneNote 一括エクスポート用 PowerShell スクリプト `export_onenote.ps1`
  （差分エクスポート方式、`_docgrep_meta.json` ベース）
- OneNote 用 PS1 を BOM 付き UTF-8 + chcp 65001 でセットアップして
  Windows PowerShell 5.x の文字化けを抑止
- `runtime.parallel` 並列実行（COM は直列維持、text/xlsx を ThreadPoolExecutor）
- `extractors/text.py` に UTF-8 高速パスとバイナリ判定（先頭 8KB + BOM チェック）
- 拡張子限定の解除 (`extensions: ["*"]`) と既知バイナリ拡張子の即スキップ
- 出力先既定を `reports/` に変更、HTML の `latest_path` ミラー機能
- スキップ理由ごとのパス記録 + HTML モーダル表示 + Excel `skipped` シート
- 出力時の制御文字サニタイズ (`openpyxl.ILLEGAL_CHARACTERS_RE`)
- openpyxl の `Cannot parse header or footer` / `extension is not supported`
  系警告の抑制
