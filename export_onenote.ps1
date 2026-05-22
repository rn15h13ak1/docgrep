<#
.SYNOPSIS
    OneNote の全ノートブックを Word(.docx) に一括エクスポートする（docgrep 検索用）。

.DESCRIPTION
    OneNote に現在開かれている（登録済みの）全ノートブックを COM 経由で取得し、
    セクション単位またはページ単位で Word(.docx) として出力します。
    出力先は docgrep の検索対象に含まれる専用フォルダで、実行のたびに中身をクリアして再生成します。

    Python(pywin32) からの COM はこの環境では「ライブラリは登録されていません」で失敗するため、
    PowerShell から COM を呼ぶ構成にしています（実機検証済み）。

.NOTES
    - 管理者ではない「通常権限」の PowerShell で実行してください
      （管理者として実行すると、通常権限で動作する OneNote と別セッションになり RPC で繋がらないことがあります）。
    - 実行中は OneNote を起動したまま閉じないでください
      （途中で開閉すると COM 接続が切れて RPC エラー 0x800706BE / 0x800706BA が発生します）。
    - 手書き（インク）・画像内の文字は Word 出力されません（COMエクスポート共通の制約）。

.PARAMETER Granularity
    エクスポートの粒度。'section'（既定, セクション単位で1ファイル）または 'page'（ページ単位で1ファイル）。

.PARAMETER OutDir
    出力先フォルダ。既定はスクリプトと同じ場所の .\onenote_export 。
    docgrep の検索対象パスに含まれている必要があります。

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\export_onenote.ps1
    powershell -ExecutionPolicy Bypass -File .\export_onenote.ps1 -Granularity page
#>

[CmdletBinding()]
param(
    [ValidateSet('section', 'page')]
    [string]$Granularity = 'section',

    [string]$OutDir = (Join-Path $PSScriptRoot 'onenote_export')
)

# --- 文字化け対策（Windows PowerShell 5.x 向け） ---
# ファイル自体は BOM 付き UTF-8 で保存しているが、コンソール出力エンコーディングも
# UTF-8 に揃えないと Write-Host / Write-Warning の日本語が化ける。
try {
    $null = & chcp 65001
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    $OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
    # 一部の制限された環境では失敗することがあるが、致命的ではないため握りつぶす
}

# OneNote COM の定数
$hsPages = 4          # GetHierarchy のスコープ: ノートブック→セクション→ページまで
$pfWord  = 5          # Publish のフォーマット: Word(.docx)
$xmlns   = 'http://schemas.microsoft.com/office/onenote/2013/onenote'

# ファイル名に使えない文字をアンダースコアに置換
function Sanitize-Name([string]$name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return '_' }
    return ($name -replace '[\\/:*?"<>|]', '_').Trim()
}

# COM 呼び出しのリトライ（RPCの一時的失敗に備える）
function Invoke-WithRetry([scriptblock]$Action, [int]$MaxRetry = 3, [int]$DelaySec = 2) {
    for ($i = 1; $i -le $MaxRetry; $i++) {
        try {
            return & $Action
        } catch {
            if ($i -eq $MaxRetry) { throw }
            Write-Warning "COM 呼び出しに失敗（$i/$MaxRetry）。$DelaySec 秒後に再試行します: $($_.Exception.Message)"
            Start-Sleep -Seconds $DelaySec
        }
    }
}

Write-Host '=== OneNote 一括エクスポート開始 ===' -ForegroundColor Cyan
Write-Host "粒度: $Granularity / 出力先: $OutDir"

# 1) OneNote COM へ接続
try {
    $onenote = New-Object -ComObject OneNote.Application
} catch {
    Write-Error 'OneNote COM への接続に失敗しました。OneNote を起動してから、通常権限の PowerShell で実行してください。'
    exit 1
}

# 2) 階層（開いている全ノートブック→セクション→ページ）を取得
$hierarchyXml = ''
try {
    Invoke-WithRetry { $onenote.GetHierarchy('', $hsPages, [ref]$hierarchyXml) } | Out-Null
} catch {
    Write-Error "階層の取得に失敗しました（OneNote が起動・応答しているか確認してください）: $($_.Exception.Message)"
    exit 1
}

[xml]$hierarchy = $hierarchyXml
$ns = New-Object System.Xml.XmlNamespaceManager($hierarchy.NameTable)
$ns.AddNamespace('one', $xmlns)

# 3) 出力先フォルダを毎回クリアして作り直す
if (Test-Path $OutDir) {
    Write-Host '既存の出力フォルダをクリアします...'
    Remove-Item -Path (Join-Path $OutDir '*') -Recurse -Force -ErrorAction SilentlyContinue
} else {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

# 4) エクスポート対象を粒度に応じて収集
#    name 属性は階層をたどって衝突しないファイル名を作るために使用
$targets = @()

foreach ($notebook in $hierarchy.SelectNodes('//one:Notebook', $ns)) {
    $nbName = Sanitize-Name $notebook.name
    foreach ($section in $notebook.SelectNodes('.//one:Section', $ns)) {
        $secName = Sanitize-Name $section.name
        if ($Granularity -eq 'section') {
            $targets += [pscustomobject]@{
                ID       = $section.ID
                FileName = "${nbName}__${secName}.docx"
            }
        } else {
            # page 粒度: セクション配下の各ページ
            foreach ($page in $section.SelectNodes('.//one:Page', $ns)) {
                $pgName = Sanitize-Name $page.name
                $targets += [pscustomobject]@{
                    ID       = $page.ID
                    FileName = "${nbName}__${secName}__${pgName}.docx"
                }
            }
        }
    }
}

if ($targets.Count -eq 0) {
    Write-Warning 'エクスポート対象が見つかりませんでした。OneNote に検索対象のノートブックが開かれているか確認してください。'
    exit 0
}

Write-Host "エクスポート対象: $($targets.Count) 件"

# 5) 重複ファイル名を回避しつつ Publish で .docx 出力
$ok = 0; $ng = 0
$seen = @{}
foreach ($t in $targets) {
    $fileName = $t.FileName
    if ($seen.ContainsKey($fileName)) {
        $seen[$fileName]++
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + "_$($seen[$fileName]).docx"
    } else {
        $seen[$fileName] = 0
    }
    $dest = Join-Path $OutDir $fileName

    try {
        Invoke-WithRetry { $onenote.Publish($t.ID, $dest, $pfWord, '') } | Out-Null
        if (Test-Path $dest) {
            $ok++
            Write-Host "  [OK] $fileName"
        } else {
            $ng++
            Write-Warning "  [NG] $fileName （ファイルが生成されませんでした）"
        }
    } catch {
        $ng++
        Write-Warning "  [NG] $fileName : $($_.Exception.Message)"
    }
}

Write-Host '=== 完了 ===' -ForegroundColor Cyan
Write-Host "成功: $ok 件 / 失敗: $ng 件 / 出力先: $OutDir"
if ($ng -gt 0) {
    Write-Host '失敗がある場合は、OneNote を起動したまま・通常権限で再実行してください。' -ForegroundColor Yellow
}
