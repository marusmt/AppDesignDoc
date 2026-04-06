<#
.SYNOPSIS
    CRUD解析ツール メイン実行スクリプト

.DESCRIPTION
    Oracle SQLソースとVB.NET DACソースを解析し、
    テーブル×項目×機能のCRUDマトリックスをExcelで出力する

.PARAMETER ConfigPath
    設定ファイルパス（デフォルト: .\config.psd1）

.PARAMETER ExportMode
    Excel出力方式（"Module" = ImportExcel, "COM" = Excel COM）
    デフォルト: Module

.PARAMETER SkipOracle
    Oracle解析をスキップ

.PARAMETER SkipVbNet
    VB.NET解析をスキップ

.PARAMETER SkipDdl
    DDL（テーブル定義・インデックス定義）解析をスキップ

.PARAMETER DebugOracle
    Oracle 解析でプロシージャ単位の抽出件数・除外理由のデバッグ出力を行う

.EXAMPLE
    .\Run-CrudAnalysis.ps1
    .\Run-CrudAnalysis.ps1 -ConfigPath ".\my_config.psd1" -ExportMode COM
    .\Run-CrudAnalysis.ps1 -SkipOracle
    .\Run-CrudAnalysis.ps1 -SkipDdl
#>

param(
    [string]$ConfigPath = ".\config.psd1",
    [ValidateSet("Module", "COM")]
    [string]$ExportMode = "Module",
    [switch]$SkipOracle,
    [switch]$SkipVbNet,
    [switch]$SkipDdl,
    [switch]$DebugOracle
)

$ErrorActionPreference = "Stop"
$scriptDir = $PSScriptRoot

# 関連スクリプトの読み込み
. "$scriptDir\Parse-OracleSql.ps1"
. "$scriptDir\Parse-VbNetDac.ps1"
. "$scriptDir\Parse-TableDdl.ps1"
. "$scriptDir\Export-CrudExcel.ps1"

# 設定ファイル読み込み
Write-Host "============================================" -ForegroundColor White
Write-Host "  CRUD解析ツール" -ForegroundColor White
Write-Host "============================================" -ForegroundColor White
Write-Host ""

$configFullPath = if ([System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath
} else {
    Join-Path $scriptDir $ConfigPath
}

if (-not (Test-Path $configFullPath)) {
    Write-Error "設定ファイルが見つかりません: $configFullPath"
    exit 1
}

$config = Import-PowerShellDataFile $configFullPath
Write-Host "[設定] 設定ファイル読み込み完了: $configFullPath" -ForegroundColor Green

# 解析結果格納用
$allResults = [System.Collections.ArrayList]::new()
$tableDefs = [System.Collections.ArrayList]::new()
$indexDefs = [System.Collections.ArrayList]::new()
$unusedColumns = [System.Collections.ArrayList]::new()

$oracleKnownCte = @()
if ($null -ne $config.Oracle.KnownCteNames -and $config.Oracle.KnownCteNames.Count -gt 0) {
    $oracleKnownCte = @($config.Oracle.KnownCteNames)
}

# --- Oracle SQL 解析 ---
if (-not $SkipOracle) {
    Write-Host ""
    Write-Host "--- Oracle SQL 解析 ---" -ForegroundColor Yellow

    if (-not (Test-Path $config.Oracle.SourcePath)) {
        Write-Warning "[Oracle] ソースディレクトリが見つかりません: $($config.Oracle.SourcePath)"
        Write-Warning "[Oracle] Oracle解析をスキップします"
    }
    else {
        $oracleParams = @{
            SourcePath       = $config.Oracle.SourcePath
            FilePattern      = $config.Oracle.FilePattern
            ExcludePatterns  = $config.Oracle.ExcludePatterns
            ExcludeTables    = $config.ExcludeTables
            ExcludeSchemas   = $config.ExcludeSchemas
            AdditionalCteNames = $oracleKnownCte
        }
        if ($DebugOracle) { $oracleParams.DebugLog = $true }
        elseif ($null -ne $config.Oracle.DebugLog -and $config.Oracle.DebugLog -eq $true) { $oracleParams.DebugLog = $true }
        $oracleResults = ConvertFrom-OracleSqlDirectory @oracleParams

        foreach ($r in $oracleResults) {
            [void]$allResults.Add($r)
        }
    }
}
else {
    Write-Host "[Oracle] スキップ" -ForegroundColor DarkGray
}

# --- VB.NET DAC 解析 ---
if (-not $SkipVbNet) {
    Write-Host ""
    Write-Host "--- VB.NET DAC 解析 ---" -ForegroundColor Yellow

    if (-not (Test-Path $config.VbNet.SourcePath)) {
        Write-Warning "[VB.NET] ソースディレクトリが見つかりません: $($config.VbNet.SourcePath)"
        Write-Warning "[VB.NET] VB.NET解析をスキップします"
    }
    else {
        $vbnetResults = ConvertFrom-VbNetDacDirectory `
            -SourcePath $config.VbNet.SourcePath `
            -DacFilePattern $config.VbNet.DacFilePattern `
            -ExcludePatterns $config.VbNet.ExcludePatterns `
            -ExcludeTables $config.ExcludeTables `
            -ExcludeSchemas $config.ExcludeSchemas `
            -KnownCteNames $oracleKnownCte

        foreach ($r in $vbnetResults) {
            [void]$allResults.Add($r)
        }
    }
}
else {
    Write-Host "[VB.NET] スキップ" -ForegroundColor DarkGray
}

# --- DDL（テーブル定義・インデックス定義）解析 ---
if (-not $SkipDdl) {
    Write-Host ""
    Write-Host "--- DDL 解析 ---" -ForegroundColor Yellow

    # テーブル定義解析
    if ($config.Ddl -and $config.Ddl.TableSourcePath) {
        if (-not (Test-Path $config.Ddl.TableSourcePath)) {
            Write-Warning "[DDL] テーブル定義ディレクトリが見つかりません: $($config.Ddl.TableSourcePath)"
        }
        else {
            $ddlResult = Parse-TableDdlDirectory `
                -SourcePath $config.Ddl.TableSourcePath `
                -FilePattern $config.Ddl.FilePattern `
                -ExcludePatterns $config.Ddl.ExcludePatterns `
                -ExcludeTables $config.ExcludeTables

            foreach ($def in $ddlResult.TableDefinitions) {
                [void]$tableDefs.Add($def)
            }
            foreach ($def in $ddlResult.IndexDefinitions) {
                [void]$indexDefs.Add($def)
            }
        }
    }

    # インデックス定義解析（テーブルと別ディレクトリの場合）
    if ($config.Ddl -and $config.Ddl.IndexSourcePath -and $config.Ddl.IndexSourcePath -ne $config.Ddl.TableSourcePath) {
        if (-not (Test-Path $config.Ddl.IndexSourcePath)) {
            Write-Warning "[DDL] インデックス定義ディレクトリが見つかりません: $($config.Ddl.IndexSourcePath)"
        }
        else {
            $idxResult = Parse-TableDdlDirectory `
                -SourcePath $config.Ddl.IndexSourcePath `
                -FilePattern $config.Ddl.FilePattern `
                -ExcludePatterns $config.Ddl.ExcludePatterns `
                -ExcludeTables $config.ExcludeTables

            foreach ($def in $idxResult.IndexDefinitions) {
                [void]$indexDefs.Add($def)
            }
        }
    }

    # SELECT * 展開
    if ($config.Ddl.ExpandSelectStar -and $tableDefs.Count -gt 0 -and $allResults.Count -gt 0) {
        Write-Host ""
        Write-Host "--- SELECT * 展開 ---" -ForegroundColor Yellow
        $allResults = Expand-SelectStar -CrudResults $allResults -TableDefinitions $tableDefs
    }

    # カラム存在検証（DDL突合せ）
    if ($tableDefs.Count -gt 0 -and $allResults.Count -gt 0) {
        Write-Host ""
        Write-Host "--- カラム存在検証 ---" -ForegroundColor Yellow
        $validationResult = Test-ColumnExistence -CrudResults $allResults -TableDefinitions $tableDefs
        $allResults = $validationResult.Validated
    }

    # 未使用カラム検出
    if ($tableDefs.Count -gt 0 -and $allResults.Count -gt 0) {
        Write-Host ""
        Write-Host "--- 未使用カラム分析 ---" -ForegroundColor Yellow
        $unusedColumns = Find-UnusedColumns -TableDefinitions $tableDefs -CrudResults $allResults
    }
}
else {
    Write-Host "[DDL] スキップ" -ForegroundColor DarkGray
}

# --- 結果サマリー ---
Write-Host ""
Write-Host "--- 解析結果サマリー ---" -ForegroundColor Yellow

$tableCount = ($allResults | ForEach-Object { $_.TableName } | Sort-Object -Unique).Count
$featureCount = ($allResults | ForEach-Object { $_.FeatureName } | Sort-Object -Unique).Count
$totalEntries = $allResults.Count

$ddlTableCount = ($tableDefs | ForEach-Object { $_.TableName } | Sort-Object -Unique).Count
$ddlColumnCount = $tableDefs.Count
$ddlIndexCount = ($indexDefs | ForEach-Object { $_.IndexName } | Sort-Object -Unique).Count

Write-Host "  検出テーブル数（CRUD） : $tableCount" -ForegroundColor White
Write-Host "  検出機能数             : $featureCount" -ForegroundColor White
Write-Host "  総CRUDエントリ         : $totalEntries" -ForegroundColor White
if ($ddlTableCount -gt 0) {
    Write-Host "  テーブル定義数（DDL）  : $ddlTableCount テーブル / $ddlColumnCount カラム" -ForegroundColor White
    Write-Host "  インデックス定義数     : $ddlIndexCount" -ForegroundColor White
    Write-Host "  未使用カラム数         : $($unusedColumns.Count)" -ForegroundColor White
}

if ($allResults.Count -eq 0 -and $tableDefs.Count -eq 0 -and $indexDefs.Count -eq 0) {
    Write-Warning "解析結果が0件です。設定ファイルのパスやパターンを確認してください。"
    exit 0
}

# --- JSON出力 ---
Write-Host ""
Write-Host "--- 中間データ出力 ---" -ForegroundColor Yellow

$jsonPath = if ([System.IO.Path]::IsPathRooted($config.Output.JsonPath)) {
    $config.Output.JsonPath
} else {
    Join-Path $scriptDir $config.Output.JsonPath
}

if ($allResults.Count -gt 0) {
    Export-CrudJson -CrudResults $allResults -JsonPath $jsonPath
}
else {
    Write-Host "[JSON] CRUDエントリが0件のためスキップ" -ForegroundColor DarkGray
}

# --- Excel出力 ---
Write-Host ""
Write-Host "--- Excel出力 ---" -ForegroundColor Yellow

$excelPath = if ([System.IO.Path]::IsPathRooted($config.Output.ExcelPath)) {
    $config.Output.ExcelPath
} else {
    Join-Path $scriptDir $config.Output.ExcelPath
}

$exportParams = @{
    CrudResults      = $allResults
    ExcelPath        = $excelPath
    SummarySheetName = $config.Output.SummarySheetName
    DetailSheetName  = $config.Output.DetailSheetName
}
if ($tableDefs.Count -gt 0) { $exportParams.TableDefinitions = $tableDefs }
if ($indexDefs.Count -gt 0) { $exportParams.IndexDefinitions = $indexDefs }
if ($unusedColumns.Count -gt 0) { $exportParams.UnusedColumns = $unusedColumns }

switch ($ExportMode) {
    "Module" {
        Export-CrudExcelWithModule @exportParams
    }
    "COM" {
        Export-CrudExcelWithCOM @exportParams
    }
}

# --- 完了 ---
Write-Host ""
Write-Host "============================================" -ForegroundColor Green
Write-Host "  CRUD解析完了" -ForegroundColor Green
Write-Host "============================================" -ForegroundColor Green
Write-Host "  JSON: $jsonPath" -ForegroundColor White
Write-Host "  Excel: $excelPath" -ForegroundColor White
Write-Host ""
