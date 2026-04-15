#Requires -Version 5.1
<#
.SYNOPSIS
    プログラムソースコードからSQL文を抽出し、単独実行可能な .sql ファイルとして出力します。
.DESCRIPTION
    PL/SQL および VB.NET のソースコードを解析し、静的SQL・動的SQL・IF分岐内のSQL断片を
    すべて抽出・展開して、個別の .sql ファイルとして出力します。
.PARAMETER InputPath
    入力ファイルまたはフォルダのパス（必須）
.PARAMETER OutputDir
    出力先ディレクトリ（省略時: ./output）
.PARAMETER Language
    言語指定: "plsql" | "vbnet" | "auto"（省略時: auto）
.PARAMETER Encoding
    文字コード（省略時: UTF8）
.EXAMPLE
    .\Extract-Sql.ps1 -InputPath "./src/OrderProc.pkb"
.EXAMPLE
    .\Extract-Sql.ps1 -InputPath "./src" -OutputDir "./sql_output" -Language auto
.EXAMPLE
    .\Extract-Sql.ps1 -InputPath "./src/OrderForm.vb" -Language vbnet -Encoding UTF8
#>

param(
    [Parameter(Mandatory = $true, Position = 0,
        HelpMessage = '入力ファイルまたはフォルダのパスを指定してください')]
    [string]$InputPath,

    [Parameter()]
    [string]$OutputDir = './output',

    [Parameter()]
    [ValidateSet('plsql', 'vbnet', 'auto')]
    [string]$Language = 'auto',

    [Parameter()]
    [string]$Encoding = 'UTF8'
)

# ============================================================
# モジュールのインポート
# ============================================================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulesDir = $scriptDir

Import-Module (Join-Path $modulesDir 'SqlFormatter.psm1') -Force
Import-Module (Join-Path $modulesDir 'PlSqlParser.psm1') -Force
Import-Module (Join-Path $modulesDir 'VbNetParser.psm1') -Force

# ============================================================
# ログファイルの初期化
# ============================================================
$logDir = Join-Path $OutputDir 'logs'
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory -Force | Out-Null
}
$logFile = Join-Path $logDir ("extract-sql_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')

# ============================================================
# Get-SourceLanguage: 言語自動判定
# ============================================================
function Get-SourceLanguage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string]$ForcedLanguage = 'auto'
    )

    if ($ForcedLanguage -ne 'auto') {
        return $ForcedLanguage
    }

    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()

    switch ($ext) {
        '.sql'   { return 'plsql' }
        '.pls'   { return 'plsql' }
        '.pck'   { return 'plsql' }
        '.pkb'   { return 'plsql' }
        '.pks'   { return 'plsql' }
        '.plb'   { return 'plsql' }
        '.vb'    { return 'vbnet' }
        '.vbnet' { return 'vbnet' }
        default  { return $null }
    }
}

# ============================================================
# Get-TargetFiles: 処理対象ファイルの収集
# ============================================================
function Get-TargetFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $targetExtensions = @('.sql', '.pls', '.pck', '.pkb', '.pks', '.plb', '.vb', '.vbnet')

    if (Test-Path $Path -PathType Leaf) {
        # 単一ファイル
        return @(Get-Item $Path)
    }
    elseif (Test-Path $Path -PathType Container) {
        # フォルダ → 再帰走査
        return Get-ChildItem -Path $Path -Recurse -File |
            Where-Object { $targetExtensions -contains $_.Extension.ToLower() }
    }
    else {
        Write-Log -Level ERROR -Message "入力パスが見つかりません: $Path" -LogFile $logFile
        return @()
    }
}

# ============================================================
# メイン処理
# ============================================================
function Main {
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Green
    Write-Host '  SQL Extractor Tool v1.0' -ForegroundColor Green
    Write-Host '========================================' -ForegroundColor Green
    Write-Host ''

    Write-Log -Level INFO -Message "InputPath: $InputPath" -LogFile $logFile
    Write-Log -Level INFO -Message "OutputDir: $OutputDir" -LogFile $logFile
    Write-Log -Level INFO -Message "Language:  $Language" -LogFile $logFile
    Write-Log -Level INFO -Message "Encoding:  $Encoding" -LogFile $logFile
    Write-Host ''

    # 入力パスの検証
    if (-not (Test-Path $InputPath)) {
        Write-Log -Level ERROR -Message "入力パスが存在しません: $InputPath" -LogFile $logFile
        exit 1
    }

    # 出力ディレクトリの作成
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        Write-Log -Level INFO -Message "出力ディレクトリを作成: $OutputDir" -LogFile $logFile
    }

    # 処理対象ファイルの収集
    $targetFiles = Get-TargetFiles -Path $InputPath

    if ($targetFiles.Count -eq 0) {
        Write-Log -Level WARN -Message "処理対象のファイルが見つかりませんでした" -LogFile $logFile
        exit 0
    }

    Write-Log -Level INFO -Message "処理対象ファイル数: $($targetFiles.Count)" -LogFile $logFile
    Write-Host ''

    # 統計カウンター
    $totalFiles = 0      # 言語判定成功・解析を試みたファイル数
    $totalSkipped = 0    # 言語判定不可でスキップしたファイル数
    $totalSqls = 0
    $totalWarnings = 0
    $totalErrors = 0
    $outputFilesList = @()

    # ファイルごとに処理
    foreach ($file in $targetFiles) {
        Write-Host '----------------------------------------' -ForegroundColor DarkGray

        # 言語判定
        $lang = Get-SourceLanguage -FilePath $file.FullName -ForcedLanguage $Language

        if (-not $lang) {
            Write-Log -Level WARN -Message "言語判定不可: $($file.Name) - スキップ" -LogFile $logFile
            $totalSkipped++
            continue
        }

        $totalFiles++

        try {
            # パーサー実行
            $sqlStatements = $null

            switch ($lang) {
                'plsql' {
                    $sqlStatements = Invoke-PlSqlParser -FilePath $file.FullName -Encoding $Encoding -LogFile $logFile
                }
                'vbnet' {
                    $sqlStatements = Invoke-VbNetParser -FilePath $file.FullName -Encoding $Encoding -LogFile $logFile
                }
            }

            if ($sqlStatements -and $sqlStatements.Count -gt 0) {
                # SQL文をファイルに出力
                $outputFiles = Export-SqlFiles `
                    -SqlStatements $sqlStatements `
                    -SourceFileName $file.Name `
                    -OutputDir $OutputDir `
                    -Encoding $Encoding `
                    -LogFile $logFile

                $totalSqls += $sqlStatements.Count
                $outputFilesList += $outputFiles
            }
            else {
                Write-Log -Level INFO -Message "SQL文が見つかりませんでした: $($file.Name)" -LogFile $logFile
            }
        }
        catch {
            Write-Log -Level ERROR -Message "処理エラー: $($file.Name) - $($_.Exception.Message)" -LogFile $logFile
            $totalErrors++
        }
    }

    # ============================================================
    # サマリー出力
    # ============================================================
    Write-Host ''
    Write-Host '========================================' -ForegroundColor Green
    Write-Host '  処理完了サマリー' -ForegroundColor Green
    Write-Host '========================================' -ForegroundColor Green
    Write-Log -Level INFO -Message "Summary: $totalFiles files processed, $totalSkipped skipped, $totalSqls SQLs extracted, $totalWarnings warnings, $totalErrors errors" -LogFile $logFile
    Write-Host ''
    Write-Host "  解析ファイル数:  $totalFiles" -ForegroundColor White
    Write-Host "  スキップ数:      $totalSkipped" -ForegroundColor $(if ($totalSkipped -gt 0) { 'Yellow' } else { 'White' })
    Write-Host "  抽出SQL数:       $totalSqls" -ForegroundColor White
    Write-Host "  警告数:          $totalWarnings" -ForegroundColor $(if ($totalWarnings -gt 0) { 'Yellow' } else { 'White' })
    Write-Host "  エラー数:        $totalErrors" -ForegroundColor $(if ($totalErrors -gt 0) { 'Red' } else { 'White' })
    Write-Host "  出力先:          $OutputDir" -ForegroundColor White
    Write-Host "  ログファイル:    $logFile" -ForegroundColor White
    Write-Host ''

    if ($totalSqls -gt 0) {
        Write-Host '  出力ファイル一覧:' -ForegroundColor Cyan
        foreach ($outFile in $outputFilesList) {
            Write-Host "    - $outFile" -ForegroundColor Gray
        }
        Write-Host ''
    }

    # 終了コード
    if ($totalErrors -gt 0) {
        exit 2
    }
    exit 0
}

# エントリポイント
Main