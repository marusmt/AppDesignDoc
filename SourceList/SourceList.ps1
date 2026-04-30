<#
.SYNOPSIS
    VB.NET / PL/SQL ソースファイルを再帰走査し、言語別ファイル一覧を生成する。
.DESCRIPTION
    指定フォルダ配下を BFS で再帰走査し、拡張子に基づいて言語種別
    （VBNET / PLSQL / OTHER / SKIPPED）を判定する。
    走査結果のサマリをコンソールとログファイルに出力する。
    走査結果を UTF-8 BOM 付き CSV ファイルとして output\ に出力する。

    【実行前提】リポジトリルートをカレントディレクトリとして実行すること。
    .\config\settings.psd1 / .\output / .\logs はカレントディレクトリ相対で解決される。
.PARAMETER RootPath
    走査対象ルートフォルダのパス（必須）。存在しないパスを指定するとエラー終了する。
.PARAMETER ConfigPath
    設定ファイルのパス。v01 では未使用（指定しても WARN ログに記録して無視）。v10 で対応予定。
.PARAMETER OutputDir
    CSV 出力先フォルダのパス。既定値は .\output（カレントディレクトリ相対）。
.PARAMETER LogDir
    ログ出力先フォルダのパス。既定値は .\logs（カレントディレクトリ相対）。
.EXAMPLE
    cd C:\Repos\SourceList
    .\src\SourceList.ps1 -RootPath C:\Projects\MyApp
.NOTES
    PowerShell 5.1 専用。PS7 系では動作未確認。
    .\config\settings.psd1 が存在しない場合はハードコード定数で動作する（WARN ログ出力）。
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$RootPath,
    [string]$ConfigPath = '.\config\settings.psd1',
    [string]$OutputDir  = '.\output',
    [string]$LogDir     = '.\logs'
)

$ErrorActionPreference = 'Stop'

# ── ハードコード定数（settings.psd1 不在時のフォールバック）─────────────────────
$Script:DefaultVbNetExtensions = @(
    '.vb', '.vbproj', '.sln', '.resx', '.config', '.xml',
    '.xsd', '.xsl', '.xslt', '.txt', '.ini', '.bat', '.cmd',
    '.ps1', '.psm1', '.psd1', '.md', '.csv'
)
$Script:DefaultPlSqlExtensions = @(
    '.sql', '.pkb', '.pks', '.prc', '.fnc', '.trg', '.vw'
)
$Script:DefaultExcludeExtensions = @(
    '.exe', '.dll', '.pdb', '.lib', '.obj', '.bin', '.dat',
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', '.ico', '.tif', '.tiff',
    '.zip', '.7z', '.rar', '.cab', '.msi', '.gz', '.tar',
    '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'
)
$Script:DefaultExcludeFolders = @(
    'bin', 'obj', '.git', '.vs', 'node_modules', 'output', 'logs', '.steering'
)

# ── ① パラメータ検証（Logger 未ロード → コンソール出力のみ）────────────────────
if (-not (Test-Path -Path $RootPath -PathType Container)) {
    $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    Write-Host "$ts [ERROR] RootPath が存在しないか、フォルダではありません: $RootPath"
    exit 1
}

# ── ②〜⑧ メイン処理 ───────────────────────────────────────────────────────────
$loggerLoaded = $false
$logPath = $null
try {
    # ② LogDir 作成 + ログファイルパス決定 + Import-Module Logger
    New-Item -ItemType Directory -Path $LogDir -Force | Out-Null
    $timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
    $logFileName = "SourceList_$timestamp.log"
    $logPath     = Join-Path $LogDir $logFileName
    $csvPath     = Join-Path $OutputDir "SourceList_$timestamp.csv"
    Import-Module (Join-Path $PSScriptRoot 'Logger.psm1') -Force
    $loggerLoaded = $true

    Write-LogMessage -Level 'INFO' -Message 'SourceList v0.5 開始' -LogPath $logPath

    # ③ 設定ファイル読み込み
    if ($PSBoundParameters.ContainsKey('ConfigPath')) {
        Write-LogMessage -Level 'WARN' -Message "-ConfigPath は v01 では未対応のため無視します: $ConfigPath" -LogPath $logPath
    }
    $settingsPath = '.\config\settings.psd1'   # v01: 常に既定パスを使用
    $settings = $null
    if (Test-Path -Path $settingsPath -PathType Leaf) {
        $settings = Import-PowerShellDataFile -Path $settingsPath
        Write-LogMessage -Level 'INFO' -Message "設定ファイルを読み込みました: $settingsPath" -LogPath $logPath
    } else {
        Write-LogMessage -Level 'WARN' -Message "settings.psd1 が見つかりません。ハードコード定数で続行します: $settingsPath" -LogPath $logPath
    }

    $vbNetExtensions   = if ($settings -and $settings.VbNetExtensions)   { $settings.VbNetExtensions }   else { $Script:DefaultVbNetExtensions }
    $plSqlExtensions   = if ($settings -and $settings.PlSqlExtensions)   { $settings.PlSqlExtensions }   else { $Script:DefaultPlSqlExtensions }
    $excludeExtensions = if ($settings -and $settings.ExcludeExtensions) { $settings.ExcludeExtensions } else { $Script:DefaultExcludeExtensions }
    $excludeFolders    = if ($settings -and $settings.ExcludeFolders)    { $settings.ExcludeFolders }    else { $Script:DefaultExcludeFolders }

    # ④ OutputDir 作成
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-LogMessage -Level 'INFO' -Message "出力ディレクトリを確認しました: $OutputDir" -LogPath $logPath

    # ⑤ Import-Module（残4モジュール）
    # FileScanner は T4 で実装。T3 単独の動作確認は ④ までを範囲とし、T7 で統合検証する。
    Import-Module (Join-Path $PSScriptRoot 'FileScanner.psm1')      -Force
    Import-Module (Join-Path $PSScriptRoot 'EncodingDetector.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot 'LineCounter.psm1')      -Force
    Import-Module (Join-Path $PSScriptRoot 'CsvWriter.psm1')        -Force

    # ⑥ Get-SourceFiles 呼び出し（LogCallback 渡し）
    Write-LogMessage -Level 'INFO' -Message "走査を開始します。RootPath=$RootPath" -LogPath $logPath
    $logCallback = {
        param($Level, $Message)
        Write-LogMessage -Level $Level -Message $Message -LogPath $logPath
    }.GetNewClosure()

    $sourceFiles = Get-SourceFiles `
        -RootPath          $RootPath `
        -ExcludeFolders    $excludeFolders `
        -ExcludeExtensions $excludeExtensions `
        -PlSqlExtensions   $plSqlExtensions `
        -VbNetExtensions   $vbNetExtensions `
        -LogCallback       $logCallback

    # ⑥-b 行数カウント
    $total = $sourceFiles.Count
    Write-LogMessage -Level 'INFO' -Message "行数カウントを開始します。対象=$total 件" -LogPath $logPath
    $countIndex = 0
    foreach ($file in $sourceFiles) {
        $countIndex++
        if ($countIndex % 500 -eq 0) {
            Write-LogMessage -Level 'INFO' -Message "行数カウント進捗: $countIndex / $total 件" -LogPath $logPath
        }

        if ($file.Language -eq 'SKIPPED') {
            $enc    = 'N/A'
            $lines  = 0; $comment = 0; $blank = 0
        } else {
            $enc = Get-FileEncoding -FilePath $file.FullPath
            try {
                $counts  = Get-LineCounts -FilePath $file.FullPath -Language $file.Language -Encoding $enc
                $lines   = $counts.Lines
                $comment = $counts.CommentLines
                $blank   = $counts.BlankLines
                if ($lines -lt 0) {
                    & $logCallback 'SKIP' "（ReadError）行数カウント失敗: $($file.FullPath)"
                }
            } catch {
                $lines = -1; $comment = 0; $blank = -1
                & $logCallback 'SKIP' "（ReadError）行数カウント失敗: $($file.FullPath) : $($_.Exception.Message)"
            }
        }

        $file | Add-Member -NotePropertyName 'Encoding'     -NotePropertyValue $enc     -Force
        $file | Add-Member -NotePropertyName 'Lines'        -NotePropertyValue $lines   -Force
        $file | Add-Member -NotePropertyName 'CommentLines' -NotePropertyValue $comment -Force
        $file | Add-Member -NotePropertyName 'BlankLines'   -NotePropertyValue $blank   -Force
    }
    Write-LogMessage -Level 'INFO' -Message '行数カウント完了' -LogPath $logPath

    # ⑦ サマリ出力
    $vbnet   = @($sourceFiles | Where-Object { $_.Language -eq 'VBNET'   }).Count
    $plsql   = @($sourceFiles | Where-Object { $_.Language -eq 'PLSQL'   }).Count
    $other   = @($sourceFiles | Where-Object { $_.Language -eq 'OTHER'   }).Count
    $skipped = @($sourceFiles | Where-Object { $_.Language -eq 'SKIPPED' }).Count

    $sumLines = 0; $sumVbnet = 0; $sumPlsql = 0
    foreach ($file in $sourceFiles) {
        if ($file.Lines -ge 0) {
            $sumLines += $file.Lines
            if ($file.Language -eq 'VBNET') { $sumVbnet += $file.Lines }
            if ($file.Language -eq 'PLSQL') { $sumPlsql += $file.Lines }
        }
    }

    Write-LogMessage -Level 'INFO' -Message (
        "走査完了。総件数=$total  VBNET=$vbnet  PLSQL=$plsql  OTHER=$other  SKIPPED=$skipped  " +
        "総行数=$sumLines  VBNETの行数=$sumVbnet  PLSQLの行数=$sumPlsql"
    ) -LogPath $logPath

    # ⑦-b CSV 出力
    Export-SourceListCsv -Records $sourceFiles -OutputPath $csvPath -LogCallback $logCallback

    # ⑧ 正常終了
    exit 0

} catch {
    $errMsg = $_.Exception.Message
    if ($logPath -and $loggerLoaded) {
        Write-LogMessage -Level 'ERROR' -Message $errMsg -LogPath $logPath
    } else {
        $ts = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Write-Host "$ts [ERROR] $errMsg"
    }
    exit 1
}
