# docgrep

ファイルサーバ上の Office 文書・テキストファイルを横断検索する CLI ツール。

本文（セル値・図形内テキスト・コメント・スライド・ノート・ヘッダ/フッタ等）を抽出し、
キーワード / 正規表現 / あいまい検索を行い、Excel / HTML / コンソールに結果を出力します。

## 主要機能

- **形式横断検索**: `.xlsx` / `.docx` / `.pptx` / 旧 Office / テキスト系（行単位 + 拡張子非依存のバイナリ判定）
- **検知箇所付き結果**: `Sheet1!B5` / `スライド 5 / Title 1` / `行 12` / `コメント (Yamada)` 等の locator をヒットに紐付け
- **3 種類の検索モード**: keyword(AND/OR) / regex / fuzzy(difflib)
- **対話メニュー (`menu.py`)**: 検索 (Python) と OneNote エクスポート (PowerShell) を単一エントリで起動
- **SQLite キャッシュ + 並列処理 + per-file タイムアウト**で大量ファイルにも対応

> **動作要件が特殊です。** Anaconda（フル版）+ Windows + MS Office 前提。詳細は [動作要件](#動作要件) を必読。

---

## 目次

**入門**
1. [クイックスタート](#クイックスタート)
2. [動作要件](#動作要件)

**使い方ガイド**
3. [検索を実行する](#1-検索を実行する)
4. [OneNote を検索対象に含める](#2-onenote-を検索対象に含める)
5. [出力を読む](#3-出力を読む)
6. [パフォーマンスを上げる](#4-パフォーマンスを上げる)
7. [トラブルシューティング](#5-トラブルシューティング)

**リファレンス**
8. [対応形式と抽出範囲](#対応形式と抽出範囲)
9. [CLI オプション一覧](#cli-オプション一覧)
10. [設定ファイル `config.yaml`](#設定ファイル-configyaml)
11. [検索モード](#検索モード)
12. [Exit code](#exit-code)
13. [起動時セルフチェック](#起動時セルフチェック)
14. [既知の制約](#既知の制約)
15. [ファイル構成](#ファイル構成)

**開発者向け**
16. [テスト / CI](#テスト--ci)
17. [変更履歴](#変更履歴)

---

## クイックスタート

```cmd
:: 1. Anaconda Prompt（通常権限）で展開先に移動
cd C:/tools/docgrep

:: 2. 設定ファイルを作成して検索対象パスを書く
copy config.example.yaml config.yaml
notepad config.yaml          ← paths: の "/" 始まりに書き換え

:: 3. 対話メニューで起動
python menu.py
```

メニューから「1. 全文検索を実行」を選び、パス → キーワード → モードを答えると検索が走り、
完了後に HTML レポートを開くか聞かれます。`reports/search_result_latest.html` がブラウザで開きます。

CLI 派は直接:
```cmd
python docgrep.py "東京"
python docgrep.py "東京" "プロジェクト" --operator and
python docgrep.py "kw" --dry-run        ← 走査規模だけ確認
```

詳細は [検索を実行する](#1-検索を実行する) 以降。

---

## 動作要件

本ツールは以下の環境を **前提** とします。要件を満たさない場合、起動時セルフチェックで
中止されます。何が緩和できるかは [起動時セルフチェック](#起動時セルフチェック) を参照。

### 必須

| 要件 | 詳細 |
|---|---|
| **Windows** | Windows 10 / 11（COM 自動化に必要）|
| **Microsoft Office** | Word / Excel / PowerPoint がインストール済み（2016 以降推奨）|
| **Anaconda（フル版 Distribution）** | 同梱 Python 3.11 以上。`pip install` / `conda install` での追加導入は **不要かつ非推奨** |

<details>
<summary>利用する Anaconda 同梱パッケージ一覧</summary>

`openpyxl 3.1.5`, `lxml 4.9.3`, `charset-normalizer 2.0.4`, `psutil 5.9.0`,
`tqdm 4.65.0`, `jinja2 3.1.2`, `pywin32 305`, `PyYAML 6.0`、
および Python 標準ライブラリ (`difflib`, `re`, `zipfile`, `sqlite3`, `os` 等)。

追加の `pip` / `conda` は不要。Anaconda 同梱以外のバージョンが混ざると社内の他 PC
で動かない原因になるため非推奨。
</details>

### 推奨

- **通常権限（非管理者）の Anaconda Prompt** で実行する
  → 管理者プロンプトで起動した COM クライアントは、通常権限で動作する Office と
  別 RPC セッションになり連携できないことがあります（特に OneNote）。
- 実行中は対象ファイルを Word/Excel/PowerPoint で **開いた状態にしない**
  → docgrep は不可視・専用インスタンスで読み取り専用に開きますが、ユーザーが手動で
  開いているドキュメントには触れません。

### スコープ外（初版で非対応）

| 形式 | 扱い | 理由 |
|---|---|---|
| PDF (`.pdf`) | スキップ通知 | PyMuPDF が Anaconda 非同梱・追加インストール不可 |
| OCR（スキャン PDF・画像内文字 等）| 非対応 | Tesseract 等の外部エンジン不可 |
| OneNote (`.one`) 直接 | スキップ通知 | Python から OneNote COM が「ライブラリ未登録」になる環境制約 |
| macOS / Linux | テキスト・`.xlsx` のみ動作 | COM 利用不可 |

OneNote は同梱の `export_onenote.ps1` で `.docx` に一括変換すれば検索可能（[後述](#2-onenote-を検索対象に含める)）。

---

# 使い方ガイド

## 1. 検索を実行する

起動方法は 2 通り。通常は **対話メニュー (menu.py)** を推奨します。

### 対話メニュー（推奨・単一エントリポイント）

```cmd
python menu.py
```

`menu.py` は **検索 (Python) と OneNote エクスポート (PowerShell) を内部で呼び出す**
ハブで、3 項目のメインメニューから両系統の処理を起動できます。

| 項目 | 内部で実行されるもの |
|---|---|
| **1. 全文検索を実行** | `python docgrep.py …`（引数は対話入力から自動組み立て）|
| **2. OneNote エクスポート（docgrep 用前処理）** | `powershell -ExecutionPolicy Bypass -File export_onenote.ps1` |
| **3. OneNote エクスポート → 全文検索（連続実行）** | 上記 2 → 1 を順に実行 |

検索フロー (項目 1) では以下を対話的に指定できます:

- 設定ファイル（既定の自動検出 / 別ファイル指定）
- 検索パス（config.yaml の `paths` に従う / 手動指定 / **前回のパスを再利用**）
- 検索キーワード（**前回値が `[履歴]` で提示され Enter で再利用**）
- 検索モード（**前回モードに `← 前回` バッジ**）
- モード固有オプション（keyword 複数語時は AND/OR、fuzzy はしきい値）
- 詳細オプション（任意）: 大文字小文字 / NFKC 正規化 / `--verbose`
- 検索完了後、**HTML レポートを既定ブラウザで開くか確認**（`config.output.html.latest_path` を参照）

入力履歴は `~/.docgrep_history.json` に自動保存されます。

### CLI 直叩き

自動化・スクリプト連携にはこちら:

```cmd
python docgrep.py "東京"
python docgrep.py "東京" "プロジェクト" --operator and
python docgrep.py "[0-9]{4}-[0-9]{2}-[0-9]{2}" --mode regex
python docgrep.py "プロジエクト" --mode fuzzy --fuzzy-threshold 0.7
python docgrep.py "kw" -p "//server/share/docs" --excel "reports/r_{ts}.xlsx"

python docgrep.py "kw" --dry-run        ← 走査規模だけ確認
python docgrep.py "kw" --first-hit-only ← 最初の 1 件で打ち切り
```

オプション全リストは [CLI オプション一覧](#cli-オプション一覧) 参照。

### フルパス起動と config 自動検出

スクリプトの絶対パスを別ディレクトリから指定して起動できます:

```cmd
cd C:/SomeOtherDir
python C:/tools/docgrep/docgrep.py "キーワード"
```

`config.yaml` は **CWD → スクリプト同梱** の順で検索されます。
スクリプトと一緒に動かしたい場合は `docgrep/config.yaml` を編集し、
プロジェクト固有の設定にしたい場合は CWD に `config.yaml` を置きます。

---

## 2. OneNote を検索対象に含める

Python の `pywin32` から OneNote COM を呼ぶと「ライブラリは登録されていません」
（0x8002801D）になる環境のため、**PowerShell から COM を呼んで Word(.docx) に
一括変換** する構成にしています。

### 手順（menu.py 経由・推奨）

```cmd
python menu.py
# → メニュー 2 番「OneNote エクスポート」または
#   メニュー 3 番「OneNote エクスポート → 全文検索（連続実行）」
```

`menu.py` が `powershell -ExecutionPolicy Bypass -File export_onenote.ps1` を起動します。

### 手順（PowerShell 直接実行）

1. **OneNote を起動** し、検索対象にしたいノートブックをすべて開いた状態にする
2. **通常権限の PowerShell**（管理者として実行 *しない*）を開く
3. スクリプト実行:
   ```powershell
   cd C:\tools\docgrep
   powershell -ExecutionPolicy Bypass -File .\export_onenote.ps1
   ```
   オプション:
   - `-Granularity section` （既定: セクション単位で1ファイル）
   - `-Granularity page` （ページ単位で1ファイル）
   - `-OutDir C:\path\to\dir` （出力先変更。設定の `onenote_export_dir` と一致させる）
4. 出力フォルダ `onenote_export/` に **変更のあったノートだけ** 再エクスポートされます
5. 次に docgrep を実行すると、`onenote_export/` 配下の `.docx` が走査対象に自動追加されます

### 差分更新の動作

出力フォルダにメタ情報 `_docgrep_meta.json` が保存され、次回実行時にこれを使って差分判定します。

| 状況 | 挙動 |
|---|---|
| 初回実行 | 全件エクスポート |
| ノート未変更 | スキップ（`lastModifiedTime` 一致 + ローカル .docx 存在）|
| ノート変更あり | 再エクスポート |
| ノート名変更のみ | ファイル名 rename のみ（再エクスポートしない）|
| OneNote 側で削除/移動 | ローカル .docx も削除 |
| **ごみ箱内のノート** (`isInRecycleBin="true"` 等) | **事前にスキップ**（Publish せず、エラーにもしない）|
| 粒度切り替え (`section` ↔ `page`) | 全件クリアして再生成 |
| メタファイル破損 | 警告のうえ全件再生成 |

<details>
<summary>実行ログ例</summary>

```
ごみ箱内ノート: 3 件 → スキップ
エクスポート対象: 24 件 / 前回メタ: 22 件
  [SKIP] (省略 18 件)
  [OK] ノートブック1__新規セクション.docx
  [OK] ノートブック1__更新されたセクション.docx
  [RN] 旧ノート名.docx → 新ノート名.docx
  [DEL] 削除されたセクション.docx (OneNote 側で削除/移動)
=== 完了 ===
新規/更新: 2 件 / スキップ: 21 件 / リネーム: 1 件 / 削除: 1 件 / 失敗: 0 件 / ごみ箱: 3 件
```
</details>

メタファイル `_docgrep_meta.json` は **docgrep の走査対象から自動除外** されます
（既定の `exclude.patterns` に含まれる）。

### 制約（重要）

- **実行中に OneNote を閉じない** — RPC エラー 0x800706BE / 0x800706BA の原因
- **管理者 PowerShell では動かない** — 通常権限の OneNote と別セッションになる
- **手書きインク・画像内文字は変換されない**（OneNote COM エクスポート共通の制約）
- **スクリプトは OneNote に現在開かれている全ノートブックを対象** にします。
  新しく検索対象に加えたいノートブックは事前に OneNote で開いて一覧に登録してください。
- タスクスケジューラからの起動は本ツールでは想定しません。検索前に手動で実行してください。

---

## 3. 出力を読む

検索結果は **コンソール / Excel / HTML** の 3 形態で同時出力されます。
既定の出力先は `reports/` フォルダ配下（自動作成）。

### コンソール
ヒットごとに `[locator] スニペット` 形式で表示。最後にサマリ表。

### Excel レポート

**4 シート構成**（該当データが無いシートは省略）:

| シート | 内容 |
|---|---|
| `results` | **1 ヒット 1 行**、列 = パス / 拡張子 / 検知箇所 / ヒット語 / スニペット / 最終更新日時。オートフィルタ + フリーズペイン |
| `summary` | 走査統計（件数・処理時間・キーワード・スキップ内訳）|
| `errors` | 抽出失敗・検索失敗のファイル一覧（あれば）|
| `skipped` | スキップされたファイルの **理由 + パス** 一覧（あれば、オートフィルタ付き）|

セル値の OOXML 禁止制御文字 (`\x00`-`\x08`, `\x0B`, `\x0C`, `\x0E`-`\x1F`) は
自動でサニタイズされます (`IllegalCharacterError` を防止)。

### HTML レポート

- サマリーカード + **「スキップ詳細」ボタン群** （理由ごとの `<dialog>` モーダルでパス一覧）
- **フィルタバー**（ヒットあり時）: パス・スニペット部分一致のテキスト検索 + 拡張子チップ
  + 「N / M ファイル表示」ライブカウンタ
- ヒットファイルごとにパス（`file:///` リンク）+ ヒット件数
- 各ヒット行の左に locator バッジ、右にハイライト付きスニペット
- 末尾に「エラー一覧」セクション（赤帯）

### 出力先と「最新版」HTML

```yaml
output:
  excel:
    path: "reports/search_result_{ts}.xlsx"     # 履歴付き
  html:
    path: "reports/search_result_{ts}.html"     # 履歴付き
    latest_path: "reports/search_result_latest.html"  # 同内容を固定パスにも出力
```

- `{ts}` プレースホルダーは実行時刻 `YYYYMMDD-HHMMSS` に置換 → **履歴ファイルとして残る**
- 含めなければ毎回上書き保存
- `latest_path` を指定すると **タイムスタンプ無しの最新版** が常に同じ場所に書き出される。
  ブラウザのブックマーク・他ツールからの参照リンク等の固定先用
- Excel には `latest_path` 相当はありません（履歴前提）

### レポート索引 `reports/_index.html`

検索を実行するたびに、HTML 出力先と同じディレクトリに **`_index.html` を自動生成** します。

- 過去の HTML / Excel レポートを **mtime 降順** で一覧表示
- `latest_path` のファイルがあれば別枠の「最新版」ボックスで強調
- 各エントリにファイル名・タイムスタンプ・サイズを表示

ブラウザのブックマークを `reports/_index.html` に固定しておくと、過去の検索結果を 1 ページから辿れます。

---

## 4. パフォーマンスを上げる

### 並列化（既定で有効）

text / xlsx 抽出は **ThreadPoolExecutor で並列処理** されます
（`.doc` / `.docx` / `.ppt` / `.pptx` / `.xls` は COM がスレッドセーフでないため常に直列）。

```yaml
runtime:
  parallel: 0     # 0=auto(CPU/2) / 1=直列 / N=N スレッド
```

ローカル PC への影響を抑える既定値（`CPU / 2`）から始めて、必要に応じて増やしてください。
`runtime.process_priority: below_normal` と併用するとフォアグラウンド作業を優先できます。

実行ログには次の行が出ます:
```
走査開始: 1247 ファイル (並列対象=1200 / COM 直列=47, 並列度=4) / paths=[...]
```

### SQLite 抽出キャッシュ（オプトイン）

`path + mtime + size` をキーに抽出済み Segment を SQLite に保存し、
2 回目以降の検索ではファイルを開かずに **キャッシュから直接検索** します。
同じファイル群に対して何度も検索する運用では劇的な高速化になります。

```yaml
runtime:
  cache:
    enabled: true
    path: "reports/.docgrep_cache.sqlite"
```

CLI からも `--cache` / `--no-cache` / `--cache-path` で上書き可能。
管理用コマンド: `--cache-stats` / `--cache-vacuum` / `--cache-clear`。

実行終了時に `[INFO] キャッシュ統計: hits=N, misses=N, writes=N` がログに出ます。

### xlsx Segment 粒度切替

```yaml
xlsx_granularity: cell   # または "row"
```

- `cell`（既定）: 1 セル 1 Segment、locator は `Sheet1!B5`
- `row`: 1 行 1 Segment（値はタブ結合）、locator は `Sheet1!Row 5`。
  Segment 数が列数分減って **大規模 xlsx (数十万セル) のメモリ・時間を削減**。
  セル番地は失われる代わりに高速化

### per-file タイムアウト

`runtime.per_file_timeout_sec` を秒数で指定すると、1 ファイルの抽出が指定時間を
超えたら強制打ち切り。`timeout_error: exceeded Ns` としてエラーシートに記録され、
後続ファイルは続行されます。

COM 直列フェーズでタイムアウトすると、`OfficeCom` インスタンスを破棄して
再生成可能な状態に戻します（1 つの固まった文書が後続全てを巻き込まない）。

### 走査打ち切り

```cmd
python docgrep.py "kw" --max-files 10   ← 10 件で打ち切り
python docgrep.py "kw" --first-hit-only ← 最初の 1 件で打ち切り
```

探索的検索や「とりあえずヒットファイルだけ確認したい」用途に。サマリに
`打ち切り: max_files=N に到達` として記録されます。

### テキスト抽出の UTF-8 ファストパス（自動）

実環境のテキストは大半が UTF-8 / UTF-8 BOM 付きのため、まず `utf-8-sig` で直接
デコードを試み、失敗時のみ `charset-normalizer` の文字コード推定にフォールバック
します。推定処理の重さを多くのケースで回避できます。

CP932 (Shift-JIS) / EUC-JP / UTF-16 LE/BE などは fallback 経路で正しくデコードされます。

---

## 5. トラブルシューティング

### `[NG] 必須パッケージが不足`
Anaconda フル版を使っていますか? `conda list` で `openpyxl`, `lxml`,
`charset-normalizer`, `jinja2`, `pywin32`, `pyyaml` が並んでいるか確認。
別の Python（システム Python / Miniconda / venv）が起動していないかも確認してください。

### `[NG] MS Office (COM) の起動に失敗`
- Word / Excel / PowerPoint が正常にインストールされているか
- Anaconda Prompt を **管理者として実行していない** か
- 残存プロセスがある場合は次で掃除して再試行:
  ```cmd
  taskkill /F /IM EXCEL.EXE
  taskkill /F /IM WINWORD.EXE
  taskkill /F /IM POWERPNT.EXE
  ```

### UNC パスでヒットがない / 走査されない
- `//server/share/...` 形式で書いている（`\\` ではない）
- 実行ユーザーで該当パスに **読み取り権限** がある
- `--verbose` でログを確認し、`extract_error` の中身を見る

### OneNote エクスポートで RPC エラー (`0x800706BE`)
- OneNote が起動中か
- 通常権限の PowerShell か（管理者ではない）
- スクリプト実行中に OneNote を閉じていないか
- OneNote に検索対象のノートブックを開き直して再実行

### 設定ファイルのバックスラッシュエラー
```
設定エラー: 設定ファイル ./config.yaml のパス値にバックスラッシュ (\) が含まれています。
```
YAML 内のパスはフォワードスラッシュ `/` のみ。
`\\server\share\docs` → `//server/share/docs` に置き換えてください。
詳細は [パス指定ルール](#パス指定ルール)。

### 検索が遅い
- `paths` を絞り込む / `extensions` を必要なものだけに
- `exclude.max_file_size_mb` で巨大ファイルを除外
- `runtime.process_priority: idle` で他作業を優先
- `--mode fuzzy` は遅い → `--fuzzy-threshold` を高めに
- 同じファイル群を繰り返し検索する場合は `runtime.cache.enabled: true` または `--cache`
- 大規模 xlsx が遅い → `xlsx_granularity: row` を試す
- 走査規模が読めない場合はまず `--dry-run` で対象数を確認

### menu.py 内の PowerShell 呼び出しで「'powershell' は内部コマンドまたは外部コマンドではない」
Windows 標準の PowerShell が PATH 上にあることを確認。WSL や別 OS のターミナルから
`menu.py` を叩いていないかも確認してください。

### menu.py の入力履歴をリセットしたい
`%USERPROFILE%\.docgrep_history.json`（Windows）/ `~/.docgrep_history.json`（Unix）
を削除すれば次回起動で初期化されます。

---

# リファレンス

## 対応形式と抽出範囲

| 形式 | 拡張子 | 抽出方式 | 抽出される本文 | 検知箇所 (locator) 例 |
|---|---|---|---|---|
| テキスト | 拡張子非依存（中身判定） | charset-normalizer + バイナリ判定 | 全文（行単位）| `行 12` |
| Excel 新 | `.xlsx` / `.xlsm` | openpyxl + lxml | セル値・シート名・コメント・テキストボックス・図形・SmartArt・グラフ | `Sheet1!B5` / `Sheet1: 図形「TextBox 1」` / `Sheet1!C12 コメント (Tanaka)` |
| Word | `.doc` / `.docx` | MS Office COM | 本文・ヘッダ/フッタ・図形内・コメント | `本文` / `セクション 1 / ヘッダー` / `コメント #2 (Yamada)` |
| PowerPoint | `.ppt` / `.pptx` | MS Office COM | スライド本文・図形内・ノート | `スライド 5 / Title 1` / `スライド 5 / ノート` |
| Excel 旧 | `.xls` | MS Office COM | セル値・シート名・図形 | `Sheet1!B5` / `Sheet1: 図形「Rectangle 1」` |
| OneNote (Word 化済) | `.docx` | （Word 抽出に乗る）| 同 Word | `本文` 等 |

ファイル全文を 1 つの文字列に連結するのではなく、抽出器がテキスト断片 + 出所識別子
(`locator`) の **Segment** として返し、検索結果に「どこで見つかったか」が同行表示されます。

### テキスト判定のロジック

テキスト抽出は **拡張子に依存しません**。`.json` / `.xml` / `.html` / `.ini` /
拡張子なしファイル（`README`, `Makefile` 等）でも、中身がテキストなら検索対象になります。

判定順:

1. **既知の Office 形式**（`.docx` 等）→ 専用抽出器
2. **既知のスコープ外**（`.pdf` / `.one`）→ スキップ（`scope_out`）
3. **既知のバイナリ／メディア拡張子**（`.exe` / `.zip` / `.png` / `.mp4` 等）→ 中身を見ず即スキップ（`binary`）
4. それ以外 → テキスト抽出を試行
   - 先頭 8KB を読み、BOM があれば UTF-16/UTF-8 と判定 → テキスト
   - NUL バイトを含めばバイナリと判定 → スキップ（`binary`）
   - それ以外は charset-normalizer で文字コード推定 → 行単位 Segment 化

`config.yaml` の `extensions` で対象拡張子を明示することもできます。
拡張子限定を解除したい場合は `extensions: ["*"]` を指定。

---

## CLI オプション一覧

`--help` で表示される argparse グループの順で記載。

**設定・対象パス**

| オプション | 説明 |
|---|---|
| `keywords...` | 検索キーワード（位置引数・複数可） |
| `-c CONFIG`, `--config` | 設定ファイル。未指定なら CWD → スクリプト同梱の `config.yaml` を自動検出 |
| `-p PATH`, `--path` | 検索対象パス上書き（複数指定可。設定の `paths` を置き換え）|

**検索条件**

| オプション | 説明 |
|---|---|
| `--mode {keyword,regex,fuzzy}` | 検索モード（既定: `keyword`）|
| `--operator {and,or}` | 複数キーワード時の演算子（既定: `and`）|
| `--case-sensitive` | 大文字小文字を区別 |
| `--no-normalize-width` | NFKC 正規化（全角/半角統一）を無効化 |
| `--fuzzy-threshold FLOAT` | あいまい検索のしきい値 0.0〜1.0（既定: 0.80）|
| `--snippet-chars N` | ヒット箇所前後の文字数（既定: 60）|

**出力**

| オプション | 説明 |
|---|---|
| `--excel PATH` | Excel 出力先（`{ts}` は時刻に置換）|
| `--html PATH` | HTML 出力先（`{ts}` は時刻に置換）|
| `--no-console` | コンソール出力を抑制 |

**動作制御**

| オプション | 説明 |
|---|---|
| `--no-office-check` | MS Office チェックを skip（Office 未検出環境でテキスト/.xlsx のみ検索）|
| `-v`, `--verbose` | DEBUG ログ |
| `--quiet` | 進捗バー抑制（WARNING 以上のみ）|
| `--max-files N` | N 件ヒットしたら走査打ち切り |
| `--first-hit-only` | 最初の 1 ヒットで打ち切り（`--max-files=1` 相当）|
| `--ordered-output` | 並列処理時もコンソール出力をファイル入力順に保つ |
| `--dry-run` | 走査対象の集計だけ表示して終了（実走査・抽出・検索は行わない）|

**抽出キャッシュ**

| オプション | 説明 |
|---|---|
| `--cache` / `--no-cache` | SQLite キャッシュの有効/無効（`runtime.cache.enabled` 上書き）|
| `--cache-path PATH` | キャッシュ DB のパス上書き |
| `--cache-stats` | DB サイズ / エントリ数 / 平均 Segment 数を表示して終了 |
| `--cache-vacuum` | 孤児 Segment 掃除 + SQLite VACUUM |
| `--cache-clear` | キャッシュ DB を全消去して終了 |

---

## 設定ファイル `config.yaml`

`config.example.yaml` をコピーして編集します。

### パス指定ルール

YAML 内のパスは以下の規則に従ってください。

> ⚠️ **フォワードスラッシュ `/` のみ使用可能**。バックスラッシュ `\` を含む値はロード時に
> エラー終了します（YAML のエスケープ事故防止）。

| 書き方 | OK / NG | 補足 |
|---|---|---|
| `//server/share/docs` | OK | UNC は `//` で始める |
| `Z:/projects` | OK | ドライブ + `/` |
| `./docs` | OK | 設定ファイルの親ディレクトリ基準で解決 |
| `\\server\share\docs` | **NG** | バックスラッシュ |
| `Z:\projects` | **NG** | バックスラッシュ |

**相対パスの解決**: YAML 内の相対パス（`paths` / `onenote_export_dir` /
`output.*.path` / `runtime.cache.path`）は **その config.yaml のあるディレクトリ** を
基準に絶対パス化されます。これにより `python C:/full/path/docgrep.py` のような
フルパス起動でも、CWD によらず一貫した動作になります。
CLI 引数 `--path` は手入力のため CWD 相対で解釈されます。

### 設定例（最小）

```yaml
paths:
  - "//server/share/docs"
search:
  mode: keyword
output:
  html:
    path: "reports/search_result_{ts}.html"
    latest_path: "reports/search_result_latest.html"
```

<details>
<summary>設定例（完全版）</summary>

```yaml
# 検索対象（複数指定可）
paths:
  - "//server/share/docs"      # UNC
  - "Z:/projects"               # ドライブ指定
  - "./local_docs"              # 相対パス → config.yaml の親基準

# OneNote エクスポート先（PS1 が書き出す .docx を検索対象に自動追加）
onenote_export_dir: "./onenote_export"

# 対象拡張子。["*"] で拡張子フィルタを無効化（中身判定で全ファイル試行）
extensions:
  - .txt
  - .xlsx
  - .docx
  - .pptx
  # ...

# xlsx 抽出時の Segment 粒度: "cell"（既定）または "row"
xlsx_granularity: cell

# 除外設定
exclude:
  dirs: [.git, .svn, node_modules, __pycache__]
  patterns: ["~$*", "*.tmp", ".DS_Store", "_docgrep_meta.json"]
  max_file_size_mb: 100

# 検索設定
search:
  mode: keyword          # keyword / regex / fuzzy
  operator: and          # and / or
  case_sensitive: false
  normalize_width: true
  fuzzy_threshold: 0.80
  snippet_chars: 60
  max_hits_per_file: 20

# 実行制御
runtime:
  require_office: true              # MS Office 未検出で起動中止
  parallel: 0                       # 0=auto(CPU/2), 1=直列, N=N スレッド
  process_priority: below_normal    # normal / below_normal / idle
  com_recycle_every: 30             # COM インスタンスを N 件ごとに再生成
  per_file_timeout_sec: 0           # 1 ファイルのタイムアウト秒。0=無効
  cache:
    enabled: false                  # SQLite 抽出キャッシュ
    path: "reports/.docgrep_cache.sqlite"

# 出力（{ts} は実行時刻 YYYYMMDD-HHMMSS に置換）
output:
  console: true
  excel:
    enabled: true
    path: "reports/search_result_{ts}.xlsx"
  html:
    enabled: true
    path: "reports/search_result_{ts}.html"
    latest_path: "reports/search_result_latest.html"
```
</details>

設定値は読込時に **型・列挙・範囲チェック** が走り、違反は ConfigError でまとめて報告されます。

---

## 検索モード

### `keyword`

| 演算子 | 挙動 |
|---|---|
| `and` | **ファイル全体** に全キーワードが存在 → ヒット（同一セル内に全部ある必要はない）|
| `or` | いずれかが存在 → ヒット |

- `normalize_width: true`（既定）で NFKC 正規化（`ABC` ↔ `ＡＢＣ` を同一視）
- `case_sensitive: false`（既定）で大文字小文字を無視

### `regex`

Python `re` の正規表現。`--case-sensitive` 未指定なら `IGNORECASE` フラグが自動付与。
**マルチライン正規表現は動作しません** — テキスト系は行単位で Segment 化しているため。

### `fuzzy`

標準ライブラリ `difflib.SequenceMatcher` による行ベース類似度マッチ。
タイプミス・表記ゆれの吸収に有効。`fuzzy_threshold` は 0〜1 で、既定 0.80。
速度は keyword/regex より遅いため、大規模走査では `--fuzzy-threshold` を高めに（例 0.9）。

---

## Exit code

スクリプト / バッチ / タスクスケジューラからの自動化判定に利用できます。

| code | 意味 |
|---|---|
| `0` | ヒットあり |
| `1` | ヒットなし |
| `2` | 設定エラー / セルフチェック失敗 / キーワード未指定 |
| `130` | Ctrl+C による中断 |

---

## 起動時セルフチェック

3 段階で挙動が変わります:

| 状況 | 挙動 |
|---|---|
| **必須パッケージ不足** または **MS Office 未検出** | `[NG]` を表示して exit 2（起動中止）|
| **PDF / OneNote(.one) ファイル発見** | `[WARN]` で「該当形式はスキップする」旨を表示して続行 |
| **すべて充足** | `[OK]` を表示して通常実行 |

Office 未検出環境でテキスト / `.xlsx` のみ検索したい場合は次のいずれかで緩和:
- CLI: `--no-office-check`
- 設定: `runtime.require_office: false`

---

## 既知の制約

| 制約 | 詳細 |
|---|---|
| マルチライン正規表現 | テキスト系は行単位 Segment のため `\n` をまたぐパターンはマッチしない |
| PDF / OCR / OneNote(.one直接) | 初版スコープ外（OneNote は PS1 で Word 化して回避）|
| COM 並列 | COM はスレッドセーフではないため Word/PPT/旧 Excel は常に直列。text/xlsx のみ並列化 |
| パスワード保護ファイル | 開けない場合はスキップ → エラー一覧に記録 |
| 巨大 xlsx | セル単位 (`xlsx_granularity: cell`) は数十万セルでメモリ・速度負荷大。`row` で軽量化可 |
| 画像内 / OCR | 未対応。スキャン PDF や画像中の文字は検索されない |

---

## ファイル構成

```
docgrep/
├── menu.py                   ← 対話メニュー（推奨エントリポイント。検索と OneNote 両方を起動）
├── docgrep.py                ← 検索 CLI のエントリポイント (python docgrep.py "kw")
├── cli.py                    ← 引数処理 + 走査・出力ドライバ
├── config.py / config.example.yaml
├── selfcheck.py              ← 起動時依存チェック
├── walker.py                 ← ファイル走査
├── normalize.py / snippet.py / search.py
├── priority.py / utils.py
├── cache.py                  ← SQLite 抽出キャッシュ (SegmentCache)
├── extractors/
│   ├── text.py               ← charset-normalizer（行単位 + バイナリ判定）
│   ├── xlsx.py               ← openpyxl + lxml
│   └── office_com.py         ← Word / PPT / 旧 Excel (_AppHandle ベース)
├── reporter/
│   ├── console.py            ← コンソール出力
│   ├── excel.py              ← Excel 4 シート（results/summary/errors/skipped）
│   ├── html.py               ← HTML（フィルタバー + スキップモーダル）
│   └── index.py              ← reports/_index.html 一覧画面
├── tests/                    ← pytest スイート（112 件）
├── pytest.ini
├── .github/workflows/test.yml ← Python 3.11/3.12/3.13 マトリクス CI
├── export_onenote.ps1        ← OneNote → Word 一括エクスポート（差分更新）
├── CHANGELOG.md              ← Sprint 単位の変更履歴
├── README.md
└── .gitignore
```

---

# 開発者向け

## テスト / CI

Anaconda Prompt で:

```cmd
cd C:/tools/docgrep
pytest
```

`pytest.ini` で `testpaths = tests` を指定済み。pytest は Anaconda 標準同梱です。
GitHub Actions により `main` への push / PR ごとに自動実行されます
(`.github/workflows/test.yml`、Python 3.11 / 3.12 / 3.13 マトリクス)。

<details>
<summary>主なテスト一覧（全 112 件、12 ファイル）</summary>

| ファイル | 件数 | 内容 |
|---|---|---|
| `test_search.py` | 14 | keyword(AND/OR) / regex / fuzzy / NFKC / locator 伝播 |
| `test_walker.py` | 8 | 拡張子 / 除外 / サイズ / `*` ワイルドカード / 単体ファイル |
| `test_config.py` | 10 | バックスラッシュ検証 / 相対パス解決 / ConfigError / 型検証 |
| `test_extractors_text.py` | 12 | 行番号 locator / CRLF / 空行 / UTF-8 fastpath / Shift-JIS fallback / 拡張子なし |
| `test_extractors_xlsx.py` | 8 | シート名 / セル座標 / 行粒度 / コメント (author 付) / 破損ファイル耐性 |
| `test_cache.py` | 6 | put/get round-trip / mtime 変更検知 / forget / 上書き / 永続化 |
| `test_reporters.py` | 4 | Excel/HTML round-trip / 制御文字サニタイズ / モーダル表示有無 |
| `test_excel_sanitize.py` | 4 | `_sanitize` / `_row` の単体テスト |
| `test_cli_integration.py` | 11 | `_collect_files` / `_partition_files` / `_run_scan` / `_dry_run` / `_handle_cache_command` |
| `test_menu.py` | 13 | 履歴 I/O / subprocess モック / レポートパス解決 |
| `test_report_index.py` | 5 | `_index.html` 生成 / mtime 順 / latest 別枠 |
| `test_normalize.py` / `test_snippet.py` / `test_utils.py` | 17 | 共通ユーティリティ |
</details>

## 変更履歴

Sprint 単位の変更履歴は [CHANGELOG.md](CHANGELOG.md) を参照してください。
