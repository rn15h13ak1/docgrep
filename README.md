# docgrep

ファイルサーバ上の Office 文書・テキストファイルを横断検索する CLI ツール。
本文（セル値・図形内テキスト・コメント・スライド・ノート・ヘッダ/フッタ等）を抽出し、
キーワード / 正規表現 / あいまい検索を行い、Excel / HTML / コンソールに結果を出力します。

> **動作要件が特殊です。必ず「動作要件」の章を読んでからセットアップしてください。**

---

## 目次

1. [動作要件](#動作要件)
2. [対応形式と抽出範囲](#対応形式と抽出範囲)
3. [インストール](#インストール初回セットアップ)
4. [基本的な使い方](#基本的な使い方)
5. [設定ファイル `config.yaml`](#設定ファイル-configyaml)
6. [OneNote の検索手順](#onenote-の検索手順)
7. [検索モード](#検索モード)
8. [出力](#出力)
9. [Exit code](#exit-code)
10. [起動時セルフチェック](#起動時セルフチェック)
11. [ファイル構成](#ファイル構成)
12. [トラブルシューティング](#トラブルシューティング)
13. [既知の制約](#既知の制約)
14. [テスト](#テスト)

---

## 動作要件

本ツールは以下の環境を **前提** とします。要件を満たさない場合、起動時セルフチェックで
中止されます。要件のうち何が緩和できるかは [起動時セルフチェック](#起動時セルフチェック) を参照。

### 必須

| 要件 | 詳細 |
|---|---|
| **Windows** | Windows 10 / 11（COM 自動化に必要）|
| **Microsoft Office** | Word / Excel / PowerPoint がインストール済み（2016 以降推奨）|
| **Anaconda（フル版 Distribution）** | 同梱 Python 3.11 以上。`pip install` / `conda install` での追加導入は **不要かつ非推奨** |

利用する同梱パッケージ（バージョンは検証済み）:
`openpyxl 3.1.5`, `lxml 4.9.3`, `charset-normalizer 2.0.4`, `psutil 5.9.0`,
`tqdm 4.65.0`, `jinja2 3.1.2`, `pywin32 305`, `PyYAML 6.0`、および Python 標準ライブラリ
（`difflib`, `re`, `zipfile`, `os` 等）。

### 推奨

- **通常権限（非管理者）の Anaconda Prompt** で実行する
  → 管理者プロンプトで起動した COM クライアントは、通常権限で動作する Office と
  別 RPC セッションになり連携できないことがあります（特に OneNote）。
- 実行中は Word / Excel / PowerPoint / OneNote を「対象ファイルを開いた状態」にしない
  → docgrep は不可視・専用インスタンスで読み取り専用に開きますが、ユーザーが手動で
  開いているドキュメントには触れません。

### スコープ外（初版で非対応）

| 形式 | 扱い | 理由 |
|---|---|---|
| PDF (`.pdf`) | スキップ通知 | PyMuPDF が Anaconda 非同梱・追加インストール不可 |
| OCR（スキャン PDF 等）| 非対応 | Tesseract 等の外部エンジン不可 |
| OneNote (`.one`) 直接 | スキップ通知 | Python からの OneNote COM が「ライブラリ未登録」になる環境制約 |
| macOS / Linux | テキスト・`.xlsx` のみ動作 | COM 利用不可 |

> OneNote は同梱の PowerShell スクリプト `export_onenote.ps1` で `.docx` に一括変換すれば、
> docgrep で通常検索できます。詳細は [OneNote の検索手順](#onenote-の検索手順)。

---

## 対応形式と抽出範囲

| 形式 | 拡張子 | 抽出方式 | 抽出される本文 | 検知箇所 (locator) 例 |
|---|---|---|---|---|
| テキスト | 拡張子非依存（中身判定） | charset-normalizer + バイナリ判定 | 全文（行単位）| `行 12` |
| Excel 新 | `.xlsx` / `.xlsm` | openpyxl + lxml | セル値・シート名・コメント・テキストボックス・図形・SmartArt・グラフ | `Sheet1!B5` / `Sheet1: 図形「TextBox 1」` / `Sheet1!C12 コメント (Tanaka)` |
| Word | `.doc` / `.docx` | MS Office COM | 本文・ヘッダ/フッタ・図形内・コメント | `本文` / `セクション 1 / ヘッダー` / `コメント #2 (Yamada)` |
| PowerPoint | `.ppt` / `.pptx` | MS Office COM | スライド本文・図形内・ノート | `スライド 5 / Title 1` / `スライド 5 / ノート` |
| Excel 旧 | `.xls` | MS Office COM | セル値・シート名・図形 | `Sheet1!B5` / `Sheet1: 図形「Rectangle 1」` |
| OneNote (Word 化済) | `.docx` | （Word 抽出に乗る）| 同 Word | `本文` 等 |

ファイル全文を1つの文字列に連結するのではなく、抽出器がテキスト断片 + 出所識別子 (`locator`)
の **Segment** として返し、検索結果に「どこで見つかったか」が同行表示されます。

### テキスト判定のロジック

テキスト抽出は **拡張子に依存しません**。`.json` / `.xml` / `.html` / `.ini` / 拡張子なし
ファイル（`README`, `Makefile` 等）でも、中身がテキストなら検索対象になります。

判定順:

1. **既知の Office 形式**（`.docx` 等）→ 専用抽出器
2. **既知のスコープ外**（`.pdf` / `.one`）→ スキップ（`scope_out`）
3. **既知のバイナリ／メディア拡張子**（`.exe` / `.zip` / `.png` / `.mp4` 等）→ 中身を見ず即スキップ（`binary`）
4. それ以外 → テキスト抽出を試行
   - 先頭 8KB を読み、BOM があれば UTF-16/UTF-8 と判定 → テキスト
   - NUL バイトを含めばバイナリと判定 → スキップ（`binary`）
   - それ以外は charset-normalizer で文字コード推定 → 行単位 Segment 化

`config.yaml` の `extensions` で対象拡張子を明示することもできます（既定は Office + テキスト系の
ホワイトリスト）。拡張子限定を解除したい場合は `extensions: ["*"]` を指定。

---

## インストール（初回セットアップ）

1. Anaconda フル版がインストール済みの Windows PC を用意
2. このリポジトリ（または ZIP）を任意のフォルダに展開
   - 例: `C:/tools/docgrep`
3. `config.example.yaml` をコピーして `config.yaml` を作成
   ```cmd
   copy config.example.yaml config.yaml
   ```
4. `config.yaml` の `paths` を検索対象に書き換え（[設定ファイル](#設定ファイル-configyaml) 参照）
5. Anaconda Prompt で動作確認:
   ```cmd
   cd C:/tools/docgrep
   python docgrep.py --help
   ```

**追加の `pip` / `conda` インストールは不要かつ非推奨** です。
Anaconda 同梱以外のバージョンが混ざると、社内の他 PC で動かない原因になります。

---

## 基本的な使い方

```cmd
python docgrep.py "東京"
python docgrep.py "東京" "プロジェクト" --operator and
python docgrep.py "[0-9]{4}-[0-9]{2}-[0-9]{2}" --mode regex
python docgrep.py "プロジエクト" --mode fuzzy --fuzzy-threshold 0.7
python docgrep.py "kw" -p "//server/share/docs" --excel "reports/r_{ts}.xlsx"
```

### 対話メニュー

CLI 引数を都度組み立てたくない場合は、対話メニューから起動できます:

```cmd
python menu.py
```

メニュー項目:

1. **全文検索を実行** — 検索パス（config.yaml に従う / 手動指定）→ キーワード → モードを順に入力し、内容確認後に `docgrep.py` を実行します。
2. **OneNote エクスポート（docgrep 用前処理）** — `export_onenote.ps1` を起動し、OneNote を Word(.docx) へ一括変換します。
3. **OneNote エクスポート → 全文検索（連続実行）** — エクスポート成功後にそのまま検索フローへ進みます。

### CLI オプション一覧

| オプション | 説明 |
|---|---|
| `keywords...` | 検索キーワード（位置引数・複数可） |
| `-c CONFIG`, `--config` | 設定ファイル。未指定なら CWD → スクリプト同梱の `config.yaml` を自動検出 |
| `-p PATH`, `--path` | 検索対象パス上書き（複数指定可。設定の `paths` を置き換え）|
| `--mode {keyword,regex,fuzzy}` | 検索モード（既定: `keyword`）|
| `--operator {and,or}` | 複数キーワード時の演算子（既定: `and`）|
| `--case-sensitive` | 大文字小文字を区別 |
| `--no-normalize-width` | NFKC 正規化（全角/半角統一）を無効化 |
| `--fuzzy-threshold FLOAT` | あいまい検索のしきい値 0.0〜1.0（既定: 0.80）|
| `--snippet-chars N` | ヒット箇所前後の文字数（既定: 60）|
| `--excel PATH` | Excel 出力先（`{ts}` は時刻に置換）|
| `--html PATH` | HTML 出力先（`{ts}` は時刻に置換）|
| `--no-console` | コンソール出力を抑制 |
| `--no-office-check` | MS Office チェックを skip（Office 未検出環境でテキスト/.xlsx のみ検索）|
| `-v`, `--verbose` | DEBUG ログ |
| `--quiet` | 進捗バー抑制（WARNING 以上のみ）|

### フルパス起動

スクリプトの絶対パスを別ディレクトリから指定して起動できます:

```cmd
cd C:/SomeOtherDir
python C:/tools/docgrep/docgrep.py "キーワード"
```

`config.yaml` は **CWD → スクリプト同梱** の順で検索されます。
スクリプトと一緒に動かしたい場合は `docgrep/config.yaml` を編集し、
プロジェクト固有の設定にしたい場合は CWD に `config.yaml` を置きます。

---

## 設定ファイル `config.yaml`

`config.example.yaml` をコピーして編集します。主要セクション:

```yaml
# 検索対象（複数指定可）。"/" のみ使用可。"\" を含むとエラー終了。
paths:
  - "//server/share/docs"      # UNC（スラッシュ二つで開始）
  - "Z:/projects"               # ドライブ指定（コロンの後はスラッシュ）
  - "./local_docs"              # 相対パス → この config.yaml の親ディレクトリ基準

# OneNote エクスポート先（PS1 が書き出す .docx を検索対象に自動追加）
onenote_export_dir: "./onenote_export"

# 対象拡張子（小文字）
extensions:
  - .txt
  - .xlsx
  - .docx
  - .pptx
  # ...

# 除外設定
exclude:
  dirs: [.git, .svn, node_modules, __pycache__]
  patterns: ["~$*", "*.tmp", ".DS_Store"]
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

# 出力（{ts} は実行時刻 YYYYMMDD-HHMMSS に置換、含めなければ上書き保存）
output:
  console: true
  excel:
    enabled: true
    path: "reports/search_result_{ts}.xlsx"
  html:
    enabled: true
    path: "reports/search_result_{ts}.html"
    latest_path: "reports/search_result_latest.html"   # タイムスタンプ無しの最新版
```

### パス指定の重要な制約

> ⚠️ **YAML 内のパスはフォワードスラッシュ `/` のみ使用可能。** バックスラッシュ `\` を含む
> 値はロード時にエラー終了します（YAML のエスケープ事故防止）。

| 書き方 | OK / NG | 補足 |
|---|---|---|
| `//server/share/docs` | OK | UNC は `//` で始める |
| `Z:/projects` | OK | ドライブ + `/` |
| `./docs` | OK | 設定ファイルの親ディレクトリ基準で解決 |
| `\\server\share\docs` | **NG** | バックスラッシュ |
| `Z:\projects` | **NG** | バックスラッシュ |

### 相対パスの解決ルール

- YAML 内の相対パス（`paths` / `onenote_export_dir` / `output.*.path`）は
  **その config.yaml のあるディレクトリ** を基準に絶対パス化されます
- これにより `python C:/full/path/docgrep.py` のようなフルパス起動でも、
  CWD によらず一貫した動作になります
- CLI 引数 `--path` は手入力のため CWD 相対で解釈されます

---

## OneNote の検索手順

Python の `pywin32` から OneNote COM を呼ぶと「ライブラリは登録されていません」
（0x8002801D）になる環境のため、**PowerShell から COM を呼んで Word(.docx) に
一括変換** する構成にしています。

### 手順

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
4. 出力フォルダ `onenote_export/` に **変更のあったノートだけ** 再エクスポートされます（[差分更新](#差分更新の動作)）
5. 次に docgrep を実行すると、`onenote_export/` 配下の `.docx` が走査対象に自動追加されます

### 差分更新の動作

出力フォルダにメタ情報 `_docgrep_meta.json` が保存され、次回実行時にこれを使って差分判定します。

| 状況 | 挙動 |
|---|---|
| 初回実行 | 全件エクスポート |
| ノート未変更 | スキップ（`lastModifiedTime` が一致 + ローカル .docx 存在）|
| ノート変更あり | 再エクスポート |
| ノート名変更のみ | ファイル名 rename のみ（再エクスポートしない）|
| OneNote 側で削除/移動 | ローカル .docx も削除 |
| 粒度切り替え (`section` ↔ `page`) | 全件クリアして再生成 |
| メタファイル破損 | 警告のうえ全件再生成 |

実行ログ例:
```
エクスポート対象: 24 件 / 前回メタ: 22 件
  [SKIP] (省略 18 件)
  [OK] ノートブック1__新規セクション.docx
  [OK] ノートブック1__更新されたセクション.docx
  [RN] 旧ノート名.docx → 新ノート名.docx
  [DEL] 削除されたセクション.docx (OneNote 側で削除/移動)
=== 完了 ===
新規/更新: 2 件 / スキップ: 21 件 / リネーム: 1 件 / 削除: 1 件 / 失敗: 0 件
```

メタファイル `_docgrep_meta.json` は **docgrep の走査対象から自動除外** されます
（`exclude.patterns` の既定に含む）。

### 制約（重要）

- **実行中に OneNote を閉じない** — RPC エラー 0x800706BE / 0x800706BA の原因
- **管理者 PowerShell では動かない** — 通常権限の OneNote と別セッションになる
- **手書きインク・画像内文字は変換されない**（OneNote COM エクスポート共通の制約）
- **スクリプトは OneNote に現在開かれている全ノートブックを対象** にします。
  新しく検索対象に加えたいノートブックは、事前に OneNote で開いて一覧に登録してください。
- タスクスケジューラからの起動は本ツールでは想定しません（環境制約）。検索前に手動で
  1 回実行する運用です。

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
速度は keyword/regex より遅いため、大規模走査では `--threshold` を高めに（例 0.9）。

---

## 出力

### コンソール
ヒットごとに `[locator] スニペット` 形式で表示。最後にサマリ表。

### Excel レポート
3 シート構成:
- `results`: **1ヒット1行**、列 = パス / 拡張子 / 検知箇所 / ヒット語 / スニペット / 最終更新日時。オートフィルタ有効。
- `summary`: 走査統計（件数・処理時間・キーワード等）
- `errors`: 抽出失敗・検索失敗のファイル一覧（あれば）

### HTML レポート
- サマリーカード
- ヒットファイルごとにパス（`file:///` リンク）+ ヒット件数
- 各ヒット行の左に locator バッジ、右にハイライト付きスニペット
- 末尾に「エラー一覧」セクション（あれば）

`{ts}` プレースホルダーを `path` に含めると、実行時刻 `YYYYMMDD-HHMMSS` に置換されて
履歴ファイルとして残ります。含めなければ毎回上書き保存。

### 出力先と「最新版」HTML

既定の出力先は `reports/` フォルダ配下です（フォルダは自動作成）。
```yaml
output:
  excel:
    path: "reports/search_result_{ts}.xlsx"     # 履歴付き
  html:
    path: "reports/search_result_{ts}.html"     # 履歴付き
    latest_path: "reports/search_result_latest.html"  # 同内容を固定パスにも出力
```

`latest_path` を指定すると、タイムスタンプ無しの最新版を **同じ内容で上書き出力**
します。ブラウザのブックマーク、他ツールからの参照リンク等、固定パスが必要な
ユースケース向け。`""` / `null` / 設定削除で無効化されます。Excel には同等オプションは
ありません（履歴を残す前提のため）。

---

## 性能チューニング

### 並列化（既定で有効）

text / xlsx 抽出は **ThreadPoolExecutor で並列処理** されます（COM が必要な
`.doc`/`.docx`/`.ppt`/`.pptx`/`.xls` は COM がスレッドセーフではないため常に直列）。

設定:
```yaml
runtime:
  parallel: 0     # 0=auto(CPU/2) 1=直列 N=N スレッド
```

ローカル PC への影響を抑える既定値（`CPU / 2`）から始めて、必要に応じて増やしてください。
`runtime.process_priority: below_normal` も併用するとフォアグラウンド作業を優先できます。

ログに以下が出ます:
```
走査開始: 1247 ファイル (並列対象=1200 / COM 直列=47, 並列度=4) / paths=[...]
```

### テキスト抽出の UTF-8 ファストパス（自動）

実環境のテキストは大半が UTF-8 / UTF-8 BOM 付きのため、まず `utf-8-sig` で直接
デコードを試み、失敗時のみ `charset-normalizer` の文字コード推定にフォールバック
します。これにより推定処理の重さを多くのケースで回避できます。

CP932 (Shift-JIS) / EUC-JP / UTF-16 LE/BE などは fallback 経路で正しく
デコードされます。

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

## ファイル構成

```
docgrep/
├── docgrep.py                ← エントリポイント (python docgrep.py "kw")
├── cli.py                    ← 引数処理 + 走査・出力ドライバ
├── config.py / config.example.yaml
├── selfcheck.py              ← 起動時依存チェック
├── walker.py                 ← ファイル走査
├── normalize.py / snippet.py / search.py
├── priority.py / utils.py
├── extractors/
│   ├── text.py               ← charset-normalizer（行単位）
│   ├── xlsx.py               ← openpyxl + lxml
│   └── office_com.py         ← Word / PPT / 旧 Excel
├── reporter/
│   ├── console.py / excel.py / html.py
├── tests/                    ← pytest スイート
├── pytest.ini
├── export_onenote.ps1        ← OneNote → Word 一括エクスポート
├── README.md
└── .gitignore
```

---

## トラブルシューティング

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
YAML 内のパスはフォワードスラッシュ `/` のみ。`\\server\share\docs` → `//server/share/docs` に置き換えてください。

### 検索が遅い
- `paths` を絞り込む / `extensions` を必要なものだけに
- `exclude.max_file_size_mb` で巨大ファイルを除外
- `runtime.process_priority: idle` で他作業を優先
- `--mode fuzzy` は遅い → `--fuzzy-threshold` を高めに

---

## 既知の制約

| 制約 | 詳細 |
|---|---|
| マルチライン正規表現 | テキスト系は行単位 Segment のため `\n` をまたぐパターンはマッチしない |
| PDF / OCR / OneNote(.one直接) | 初版スコープ外（OneNote は PS1 で Word 化して回避）|
| 並列処理 | 未実装。設計上「逐次」（COM が単スレッド推奨のため）|
| パスワード保護ファイル | 開けない場合はスキップ → エラー一覧に記録 |
| 巨大 xlsx | セル単位 Segment のため数十万セルでメモリ・速度に影響あり。次フェーズで領域単位への切替を検討 |

---

## テスト

Anaconda Prompt で:

```cmd
cd C:/tools/docgrep
pytest
```

`pytest.ini` で `testpaths = tests` を指定済み。pytest は Anaconda 標準同梱です。

主なテスト:
- `tests/test_search.py` — keyword(AND/OR) / regex / fuzzy / NFKC / locator 伝播
- `tests/test_walker.py` — 拡張子 / 除外 / サイズ / 単体ファイル
- `tests/test_config.py` — バックスラッシュ検証 / 相対パス解決 / ConfigError
- `tests/test_extractors_text.py` — 行番号 locator / CRLF / 空行スキップ
- `tests/test_normalize.py` / `test_snippet.py` / `test_utils.py`
