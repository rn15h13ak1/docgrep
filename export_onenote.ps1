<#
.SYNOPSIS
    OneNote の全ノートブックを Word(.docx) に一括エクスポートする（docgrep 検索用）。

.DESCRIPTION
    OneNote に現在開かれている（登録済みの）全ノートブックを COM 経由で取得し、
    セクション単位またはページ単位で Word(.docx) として出力します。

    差分更新方式: 出力フォルダにメタ情報 _docgrep_meta.json を残しておき、次回実行時に
    各ノートの lastModifiedTime を比較して、変更があったものだけ再エクスポートします。
    未変更ノートはスキップ、ノート名変更は rename のみ、OneNote 側で削除されたノートは
    ローカル .docx も削除します。粒度（section/page）を切り替えた場合のみ全件再生成します。

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

    # 既定はスクリプト位置の onenote_export/。
    # param 内で $PSScriptRoot が空のことがある（dot-source / Invoke-Expression /
    # ISE の選択実行など）ため、param 外でフォールバック付きに解決する。
    [string]$OutDir = ''
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

# --- スクリプト位置の決定（$PSScriptRoot が空のときのフォールバック） ---
# 順序: $PSScriptRoot → $PSCommandPath の親 → $MyInvocation.MyCommand.Path の親 → CWD
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    if ($PSCommandPath) {
        $scriptDir = Split-Path -Parent $PSCommandPath
    } elseif ($MyInvocation.MyCommand.Path) {
        $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    } else {
        $scriptDir = (Get-Location).Path
        Write-Warning "スクリプト位置を特定できなかったため、CWD ($scriptDir) を基準にします。"
    }
}

# OutDir が未指定なら scriptDir/onenote_export を既定とする
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path $scriptDir 'onenote_export'
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

# 3) 出力先フォルダの準備（クリアしない）+ 前回メタの読み込み
$metaPath = Join-Path $OutDir '_docgrep_meta.json'
if (-not (Test-Path $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

# 既存メタを Hashtable に展開（ID → @{ file=...; lastModifiedTime=...; exportedAt=... }）
$prevItems = @{}
$prevGranularity = $null
if (Test-Path $metaPath) {
    try {
        $rawMeta = Get-Content -Path $metaPath -Raw -Encoding UTF8
        if ($rawMeta) {
            $parsed = $rawMeta | ConvertFrom-Json
            if ($parsed.granularity) { $prevGranularity = [string]$parsed.granularity }
            if ($parsed.items) {
                foreach ($prop in $parsed.items.PSObject.Properties) {
                    $prevItems[$prop.Name] = @{
                        file             = [string]$prop.Value.file
                        lastModifiedTime = [string]$prop.Value.lastModifiedTime
                        exportedAt       = [string]$prop.Value.exportedAt
                    }
                }
            }
        }
    } catch {
        Write-Warning "メタ情報 $metaPath の読み込みに失敗（破損？）。今回は全件再生成します: $($_.Exception.Message)"
        $prevItems = @{}
        $prevGranularity = $null
    }
}

# 粒度変更時は出力フォルダの .docx を全消去（メタも破棄）して仕切り直し
if ($prevGranularity -and ($prevGranularity -ne $Granularity)) {
    Write-Host "粒度が変わりました ($prevGranularity → $Granularity)。出力をクリアして再生成します。" -ForegroundColor Yellow
    Get-ChildItem -Path $OutDir -File -Filter '*.docx' | Remove-Item -Force -ErrorAction SilentlyContinue
    $prevItems = @{}
}

# 4) エクスポート対象を粒度に応じて収集（lastModifiedTime もここで取得）
#
# OneNote のごみ箱 (Recycle Bin / Deleted Pages) 内のノートブック・セクション・
# ページは Publish しようとすると失敗するため事前にスキップする。
# 判定属性:
#   - isInRecycleBin="true" … ごみ箱に入っているアイテム自体
#   - isRecycleBin="true"   … ごみ箱セクション本体
#   - isDeletedPages="true" … 「削除済みページ」用の特殊セクション
$targets = @()
$skippedDeleted = 0

foreach ($notebook in $hierarchy.SelectNodes('//one:Notebook', $ns)) {
    if ($notebook.isInRecycleBin -eq 'true') {
        $skippedDeleted++
        continue
    }
    $nbName = Sanitize-Name $notebook.name
    foreach ($section in $notebook.SelectNodes('.//one:Section', $ns)) {
        if ($section.isInRecycleBin -eq 'true' -or
            $section.isRecycleBin   -eq 'true' -or
            $section.isDeletedPages -eq 'true') {
            $skippedDeleted++
            continue
        }
        $secName = Sanitize-Name $section.name
        if ($Granularity -eq 'section') {
            $targets += [pscustomobject]@{
                ID           = $section.ID
                FileName     = "${nbName}__${secName}.docx"
                LastModified = [string]$section.lastModifiedTime
            }
        } else {
            foreach ($page in $section.SelectNodes('.//one:Page', $ns)) {
                if ($page.isInRecycleBin -eq 'true') {
                    $skippedDeleted++
                    continue
                }
                $pgName = Sanitize-Name $page.name
                $targets += [pscustomobject]@{
                    ID           = $page.ID
                    FileName     = "${nbName}__${secName}__${pgName}.docx"
                    LastModified = [string]$page.lastModifiedTime
                }
            }
        }
    }
}

if ($skippedDeleted -gt 0) {
    Write-Host "ごみ箱内ノート: $skippedDeleted 件 → スキップ" -ForegroundColor DarkGray
}

if ($targets.Count -eq 0) {
    Write-Warning 'エクスポート対象が見つかりませんでした。OneNote に検索対象のノートブックが開かれているか確認してください。'
    exit 0
}

Write-Host "エクスポート対象: $($targets.Count) 件 / 前回メタ: $($prevItems.Count) 件"

# 5) 差分判定 → Publish or スキップ
$ok = 0; $ng = 0; $skipped = 0; $renamed = 0
$seen = @{}
$currentItems = @{}

foreach ($t in $targets) {
    # ファイル名の重複回避（同名ノートが別ノートブックにある場合等）
    $fileName = $t.FileName
    if ($seen.ContainsKey($fileName)) {
        $seen[$fileName]++
        $fileName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) + "_$($seen[$fileName]).docx"
    } else {
        $seen[$fileName] = 0
    }
    $dest = Join-Path $OutDir $fileName

    $prev = $null
    if ($prevItems.ContainsKey($t.ID)) { $prev = $prevItems[$t.ID] }

    $needsExport = $true
    if ($prev) {
        $prevFile = $prev.file
        $prevTime = $prev.lastModifiedTime
        $prevDest = if ($prevFile) { Join-Path $OutDir $prevFile } else { $null }
        $unchanged = ($prevTime -eq $t.LastModified) -and $prevDest -and (Test-Path $prevDest)

        if ($unchanged) {
            # 未変更: ファイル名が変わっていれば rename だけ、変わってなければ何もしない
            if ($prevFile -ne $fileName) {
                try {
                    if (Test-Path $dest) { Remove-Item -Path $dest -Force -ErrorAction SilentlyContinue }
                    Move-Item -Path $prevDest -Destination $dest -Force
                    $renamed++
                    Write-Host "  [RN] $prevFile → $fileName"
                } catch {
                    # rename 失敗時は再出力にフォールバック
                    $unchanged = $false
                }
            }
            if ($unchanged) {
                $needsExport = $false
                $skipped++
            }
        } elseif ($prevDest -and $prevFile -ne $fileName -and (Test-Path $prevDest)) {
            # 変更ありかつ前回と別名 → 古いファイルは削除
            Remove-Item -Path $prevDest -Force -ErrorAction SilentlyContinue
        }
    }

    if ($needsExport) {
        if (Test-Path $dest) { Remove-Item -Path $dest -Force -ErrorAction SilentlyContinue }
        try {
            Invoke-WithRetry { $onenote.Publish($t.ID, $dest, $pfWord, '') } | Out-Null
            if (Test-Path $dest) {
                $ok++
                Write-Host "  [OK] $fileName"
            } else {
                $ng++
                Write-Warning "  [NG] $fileName （ファイルが生成されませんでした）"
                continue
            }
        } catch {
            $ng++
            Write-Warning "  [NG] $fileName : $($_.Exception.Message)"
            continue
        }
    }

    $currentItems[$t.ID] = @{
        file             = $fileName
        lastModifiedTime = $t.LastModified
        exportedAt       = if ($needsExport) { (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss') } else { $prev.exportedAt }
    }
}

# 6) 旧メタにあって今回ない ID → OneNote 側で削除/移動 → ローカル .docx も削除
$deleted = 0
foreach ($oldId in $prevItems.Keys) {
    if (-not $currentItems.ContainsKey($oldId)) {
        $oldFile = $prevItems[$oldId].file
        if ($oldFile) {
            $oldPath = Join-Path $OutDir $oldFile
            if (Test-Path $oldPath) {
                Remove-Item -Path $oldPath -Force -ErrorAction SilentlyContinue
                $deleted++
                Write-Host "  [DEL] $oldFile (OneNote 側で削除/移動)"
            }
        }
    }
}

# 7) メタ情報を保存
$newMeta = [ordered]@{
    version     = 1
    granularity = $Granularity
    updatedAt   = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
    items       = $currentItems
}
try {
    $newMeta | ConvertTo-Json -Depth 6 | Set-Content -Path $metaPath -Encoding UTF8
} catch {
    Write-Warning "メタ情報の保存に失敗: $($_.Exception.Message)"
}

Write-Host ''
Write-Host '=== 完了 ===' -ForegroundColor Cyan
Write-Host "新規/更新: $ok 件 / スキップ: $skipped 件 / リネーム: $renamed 件 / 削除: $deleted 件 / 失敗: $ng 件 / ごみ箱: $skippedDeleted 件"
Write-Host "出力先: $OutDir / メタ: $metaPath"
if ($ng -gt 0) {
    Write-Host '失敗がある場合は、OneNote を起動したまま・通常権限で再実行してください。' -ForegroundColor Yellow
    Write-Host '（"削除されたページ" 系のエラーは isInRecycleBin で事前スキップしていますが、' -ForegroundColor Yellow
    Write-Host '  個別ノートが復元/移動の途中など特殊な状態だと Publish が失敗することがあります）' -ForegroundColor Yellow
}
