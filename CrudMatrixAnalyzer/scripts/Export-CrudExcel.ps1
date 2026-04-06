<#
.SYNOPSIS
    CRUD解析結果をExcelファイルに出力する

.DESCRIPTION
    解析結果をピボットし、以下のシートを持つExcelを生成する
    1. テーブル×機能サマリー: テーブル単位でCRUD操作を一覧化
    2. 項目別詳細: テーブル×項目×機能のCRUD詳細
    3. テーブル定義: テーブル・カラム・データ型一覧
    4. インデックス定義: インデックス・テーブル・カラム一覧
    5. 未使用カラム: 定義済みだがコードから参照されていないカラム
    6. 生データ: 全解析結果の一覧

    Excel出力には ImportExcel モジュール または COM オートメーションを使用する
#>

function ConvertTo-FeatureHeader {
    param([string]$FeatureName)

    if ($FeatureName -match '^([^:]+):([^.]+)\.(.+)$') {
        return "$($Matches[1]):`n$($Matches[2])`n$($Matches[3])"
    }
    elseif ($FeatureName -match '^([^:]+):(.+)$') {
        return "$($Matches[1]):`n$($Matches[2])"
    }
    return $FeatureName
}

function Get-TableCommentLookupFromDefinitions {
    param([System.Collections.ArrayList]$TableDefinitions)

    $map = @{}
    if ($null -eq $TableDefinitions -or $TableDefinitions.Count -eq 0) { return $map }
    foreach ($def in $TableDefinitions) {
        $tn = $def.TableName
        $tc = $def.TableComment
        if (-not $map.ContainsKey($tn)) {
            $map[$tn] = if ($null -ne $tc) { $tc } else { "" }
        }
        elseif (($map[$tn] -eq '' -or $null -eq $map[$tn]) -and $null -ne $tc -and $tc -ne '') {
            $map[$tn] = $tc
        }
    }
    return $map
}

function Get-ColumnCommentLookupFromDefinitions {
    param([System.Collections.ArrayList]$TableDefinitions)

    $map = @{}
    if ($null -eq $TableDefinitions -or $TableDefinitions.Count -eq 0) { return $map }
    foreach ($def in $TableDefinitions) {
        $key = "$($def.TableName)|$($def.ColumnName)"
        $cc = $def.ColumnComment
        if (-not $map.ContainsKey($key)) {
            $map[$key] = if ($null -ne $cc) { $cc } else { "" }
        }
        elseif (($map[$key] -eq '' -or $null -eq $map[$key]) -and $null -ne $cc -and $cc -ne '') {
            $map[$key] = $cc
        }
    }
    return $map
}

function Sort-StringArrayOrdinalIgnoreCase {
    param([string[]]$Items)
    if ($null -eq $Items -or $Items.Length -eq 0) {
        return @()
    }
    $a = [string[]]::new($Items.Length)
    [System.Array]::Copy($Items, $a, $Items.Length)
    [System.Array]::Sort($a, [System.StringComparer]::OrdinalIgnoreCase)
    return $a
}

function Sort-TableColumnPairKeysOrdinalIgnoreCase {
    param([string[]]$PairKeys)
    if ($null -eq $PairKeys -or $PairKeys.Length -eq 0) {
        return @()
    }
    $rawArr = [object[]]::new($PairKeys.Length)
    $kArr = [string[]]::new($PairKeys.Length)
    for ($i = 0; $i -lt $PairKeys.Length; $i++) {
        $rawArr[$i] = $PairKeys[$i]
        $parts = [string]$PairKeys[$i] -split '\|', 2
        $kArr[$i] = "{0}`0{1}" -f $parts[0].ToUpper(), $parts[1].ToUpper()
    }
    [System.Array]::Sort($kArr, $rawArr, [System.StringComparer]::OrdinalIgnoreCase)
    return [string[]]$rawArr
}

function Get-ExcelSortString {
    param($Value)
    if ($null -eq $Value) { return '' }
    try { return [string]$Value } catch { return '' }
}

function Get-ExcelSortInt32 {
    param($Value)
    if ($null -eq $Value) { return 0 }
    try { return [System.Convert]::ToInt32($Value) } catch { return 0 }
}

function Build-CrudSummaryMatrix {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [System.Collections.ArrayList]$TableDefinitions = $null
    )

    $featureSet = [System.Collections.Generic.HashSet[string]]::new()
    $tableSet = [System.Collections.Generic.HashSet[string]]::new()
    $lookup = @{}

    $total = $CrudResults.Count
    $i = 0
    foreach ($item in $CrudResults) {
        $i++
        [void]$featureSet.Add($item.FeatureName)
        [void]$tableSet.Add($item.TableName)
        $key = "$($item.TableName)|$($item.FeatureName)"
        if (-not $lookup.ContainsKey($key)) {
            $lookup[$key] = [System.Collections.Generic.HashSet[string]]::new()
        }
        [void]$lookup[$key].Add($item.Operation)
        if ($i % 10000 -eq 0) {
            Write-Host "  サマリー索引構築中: $i / $total" -ForegroundColor Gray
        }
    }

    # 機能名・テーブル名は Sort-Object 既定だと OS カルチャ依存になるため OrdinalIgnoreCase（Python の sorted(key=upper) と整合）
    $features = Sort-StringArrayOrdinalIgnoreCase @([string[]]@($featureSet))
    $tables = Sort-StringArrayOrdinalIgnoreCase @([string[]]@($tableSet))
    $tableJaMap = Get-TableCommentLookupFromDefinitions -TableDefinitions $TableDefinitions

    $headerMap = [ordered]@{}
    foreach ($f in $features) {
        $headerMap[$f] = ConvertTo-FeatureHeader $f
    }

    $matrix = [System.Collections.ArrayList]::new()
    $tableCount = 0
    $totalTables = @($tables).Count

    foreach ($table in $tables) {
        $tableCount++
        if ($tableCount % 100 -eq 0 -or $tableCount -eq $totalTables) {
            Write-Host "  サマリー行構築中: $tableCount / $totalTables" -ForegroundColor Gray
        }

        $jaName = if ($tableJaMap.ContainsKey($table)) { $tableJaMap[$table] } else { "" }
        $row = [ordered]@{ "テーブル名" = $table; "テーブル名(日本語)" = $jaName }
        foreach ($feature in $features) {
            $header = $headerMap[$feature]
            $key = "$table|$feature"
            if ($lookup.ContainsKey($key)) {
                $ops = @($lookup[$key])
                $row[$header] = ($ops | Sort-Object) -join ""
            }
            else {
                $row[$header] = "-"
            }
        }
        [void]$matrix.Add([PSCustomObject]$row)
    }

    return $matrix
}

function Build-CrudDetailMatrix {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [System.Collections.ArrayList]$TableDefinitions = $null
    )

    $featureSet = [System.Collections.Generic.HashSet[string]]::new()
    $pairSet = [System.Collections.Generic.HashSet[string]]::new()
    $lookup = @{}

    $total = $CrudResults.Count
    $i = 0
    foreach ($item in $CrudResults) {
        $i++
        [void]$featureSet.Add($item.FeatureName)
        $pairKey = "$($item.TableName)|$($item.ColumnName)"
        [void]$pairSet.Add($pairKey)
        $key = "$pairKey|$($item.FeatureName)"
        if (-not $lookup.ContainsKey($key)) {
            $lookup[$key] = [System.Collections.Generic.HashSet[string]]::new()
        }
        [void]$lookup[$key].Add($item.Operation)
        if ($i % 10000 -eq 0) {
            Write-Host "  詳細索引構築中: $i / $total" -ForegroundColor Gray
        }
    }

    $features = Sort-StringArrayOrdinalIgnoreCase @([string[]]@($featureSet))
    $tableColumnPairs = Sort-TableColumnPairKeysOrdinalIgnoreCase @([string[]]@($pairSet))
    $tableJaMap = Get-TableCommentLookupFromDefinitions -TableDefinitions $TableDefinitions
    $colJaMap = Get-ColumnCommentLookupFromDefinitions -TableDefinitions $TableDefinitions

    $headerMap = [ordered]@{}
    foreach ($f in $features) {
        $headerMap[$f] = ConvertTo-FeatureHeader $f
    }

    $matrix = [System.Collections.ArrayList]::new()
    $pairCount = 0
    $totalPairs = @($tableColumnPairs).Count

    foreach ($pair in $tableColumnPairs) {
        $pairCount++
        if ($pairCount % 500 -eq 0 -or $pairCount -eq $totalPairs) {
            Write-Host "  詳細行構築中: $pairCount / $totalPairs" -ForegroundColor Gray
        }

        $parts = $pair -split '\|'
        $table = $parts[0]
        $column = $parts[1]

        $jaTable = if ($tableJaMap.ContainsKey($table)) { $tableJaMap[$table] } else { "" }
        $ck = "$table|$column"
        $jaCol = if ($colJaMap.ContainsKey($ck)) { $colJaMap[$ck] } else { "" }

        $row = [ordered]@{
            "テーブル名"       = $table
            "テーブル名(日本語)" = $jaTable
            "項目名"           = $column
            "項目名(日本語)" = $jaCol
        }

        foreach ($feature in $features) {
            $header = $headerMap[$feature]
            $key = "$pair|$feature"
            if ($lookup.ContainsKey($key)) {
                $ops = @($lookup[$key])
                $row[$header] = ($ops | Sort-Object) -join ""
            }
            else {
                $row[$header] = "-"
            }
        }
        [void]$matrix.Add([PSCustomObject]$row)
    }

    return $matrix
}

function Build-RawDataSheet {
    param([System.Collections.ArrayList]$CrudResults)

    $rows = [System.Collections.ArrayList]::new()
    $total = $CrudResults.Count
    $i = 0

    foreach ($item in $CrudResults) {
        $i++
        [void]$rows.Add([PSCustomObject][ordered]@{
            "ソース種別"   = $item.SourceType
            "ソースファイル" = $item.SourceFile
            "オブジェクト種別" = $item.ObjectType
            "オブジェクト名"  = $item.ObjectName
            "プロシージャ/メソッド" = $item.ProcName
            "機能名"       = $item.FeatureName
            "テーブル名"   = $item.TableName
            "項目名"       = $item.ColumnName
            "操作"         = $item.Operation
        })
        if ($i % 10000 -eq 0) {
            Write-Host "  生データ構築中: $i / $total" -ForegroundColor Gray
        }
    }

    return $rows
}

function Build-TableDefSheet {
    param([System.Collections.ArrayList]$TableDefinitions)

    $rows = [System.Collections.ArrayList]::new()
    if ($null -eq $TableDefinitions -or $TableDefinitions.Count -eq 0) { return $rows }

    $defs = @($TableDefinitions)
    $n = $defs.Count
    $kArr = [string[]]::new($n)
    $rawArr = [object[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $d = $defs[$i]
        $rawArr[$i] = $d
        $tn = (Get-ExcelSortString $d.TableName).ToUpper()
        $op = Get-ExcelSortInt32 $d.OrdinalPos
        $kArr[$i] = "{0}`0{1:D12}" -f $tn, $op
    }
    [System.Array]::Sort($kArr, $rawArr, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($def in $rawArr) {
        $tJa = if ($null -ne $def.TableComment) { $def.TableComment } else { "" }
        $cJa = if ($null -ne $def.ColumnComment) { $def.ColumnComment } else { "" }
        [void]$rows.Add([PSCustomObject][ordered]@{
            "テーブル名"       = $def.TableName
            "テーブル名(日本語)" = $tJa
            "No"               = $def.OrdinalPos
            "カラム名"         = $def.ColumnName
            "カラム名(日本語)" = $cJa
            "データ型"         = $def.DataType
            "NULL許可"         = $def.Nullable
            "DEFAULT"          = $def.HasDefault
            "ソースファイル"   = $def.SourceFile
        })
    }
    return $rows
}

function Build-IndexDefSheet {
    param([System.Collections.ArrayList]$IndexDefinitions)

    $rows = [System.Collections.ArrayList]::new()
    if ($null -eq $IndexDefinitions -or $IndexDefinitions.Count -eq 0) { return $rows }

    $defs = @($IndexDefinitions)
    $n = $defs.Count
    $kArr = [string[]]::new($n)
    $rawArr = [object[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $d = $defs[$i]
        $rawArr[$i] = $d
        $tn = (Get-ExcelSortString $d.TableName).ToUpper()
        $kind = if ($null -eq $d.DefinitionKind -or $d.DefinitionKind -eq 'INDEX') { 1 } else { 0 }
        $ix = (Get-ExcelSortString $d.IndexName).ToUpper()
        $cp = Get-ExcelSortInt32 $d.ColumnPos
        $kArr[$i] = "{0}`0{1}`0{2}`0{3:D8}" -f $tn, $kind, $ix, $cp
    }
    [System.Array]::Sort($kArr, $rawArr, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($def in $rawArr) {
        $tJa = if ($null -ne $def.TableComment) { $def.TableComment } else { "" }
        $cJa = if ($null -ne $def.ColumnComment) { $def.ColumnComment } else { "" }
        $kindJa = if ($null -ne $def.DefinitionKind -and $def.DefinitionKind -eq 'PK') { "主キー" } else { "インデックス" }
        [void]$rows.Add([PSCustomObject][ordered]@{
            "テーブル名"       = $def.TableName
            "テーブル名(日本語)" = $tJa
            "定義種別"         = $kindJa
            "インデックス名"   = $def.IndexName
            "一意性"           = $def.Uniqueness
            "カラム位置"       = $def.ColumnPos
            "カラム名"         = $def.ColumnName
            "カラム名(日本語)" = $cJa
            "ソースファイル"   = $def.SourceFile
        })
    }
    return $rows
}

function Build-UnusedColumnsSheet {
    param([System.Collections.ArrayList]$UnusedColumns)

    $rows = [System.Collections.ArrayList]::new()
    if ($null -eq $UnusedColumns -or $UnusedColumns.Count -eq 0) { return $rows }

    $defs = @($UnusedColumns)
    $n = $defs.Count
    $kArr = [string[]]::new($n)
    $rawArr = [object[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) {
        $d = $defs[$i]
        $rawArr[$i] = $d
        $tn = (Get-ExcelSortString $d.TableName).ToUpper()
        $cn = (Get-ExcelSortString $d.ColumnName).ToUpper()
        $kArr[$i] = "{0}`0{1}" -f $tn, $cn
    }
    [System.Array]::Sort($kArr, $rawArr, [System.StringComparer]::OrdinalIgnoreCase)
    foreach ($item in $rawArr) {
        [void]$rows.Add([PSCustomObject][ordered]@{
            "テーブル名" = $item.TableName
            "カラム名"   = $item.ColumnName
            "データ型"   = $item.DataType
            "NULL許可"   = $item.Nullable
        })
    }
    return $rows
}

function Export-CrudExcelWithModule {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [string]$ExcelPath,
        [string]$SummarySheetName = "テーブル×機能サマリー",
        [string]$DetailSheetName = "項目別詳細",
        [string]$RawSheetName = "生データ",
        [System.Collections.ArrayList]$TableDefinitions = $null,
        [System.Collections.ArrayList]$IndexDefinitions = $null,
        [System.Collections.ArrayList]$UnusedColumns = $null
    )

    $ddlForJa = $null
    if ($null -ne $TableDefinitions -and $TableDefinitions.Count -gt 0) { $ddlForJa = $TableDefinitions }

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "[Excel] ImportExcelモジュールをインストールします..." -ForegroundColor Yellow
        Install-Module ImportExcel -Force -Scope CurrentUser
    }
    Import-Module ImportExcel

    $outputDir = [System.IO.Path]::GetDirectoryName($ExcelPath)
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    if (Test-Path $ExcelPath) { Remove-Item $ExcelPath -Force }

    $hasBaseSheet = $false
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    if ($CrudResults.Count -gt 0) {
        Write-Host "[Excel] サマリーシート作成中... ($($CrudResults.Count) 件)" -ForegroundColor Cyan
        $summaryMatrix = Build-CrudSummaryMatrix -CrudResults $CrudResults -TableDefinitions $ddlForJa
        Write-Host "[Excel] サマリー構築完了: $($summaryMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        if ($summaryMatrix.Count -gt 0) {
            $pkg = $summaryMatrix | Export-Excel -Path $ExcelPath -WorksheetName $SummarySheetName `
                -AutoSize -BoldTopRow -PassThru
            $ws = $pkg.Workbook.Worksheets[$SummarySheetName]
            $ws.Row(1).Style.WrapText = $true
            $ws.Row(1).Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
            # Python(openpyxl) と同様: スクロール領域の左上を C2（1行目と A〜B 列を固定）
            $ws.View.FreezePanes(2, 3)
            $lastRow = $ws.Dimension.End.Row
            $lastCol = $ws.Dimension.End.Column
            if ($lastCol -ge 3) {
                $dataRange = [OfficeOpenXml.ExcelAddress]::new(2, 3, $lastRow, $lastCol)
                $cfC = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfC.Text = "C"; $cfC.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGreen; $cfC.Style.Font.Color.Color = [System.Drawing.Color]::DarkGreen
                $cfR = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfR.Text = "R"; $cfR.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightBlue; $cfR.Style.Font.Color.Color = [System.Drawing.Color]::DarkBlue
                $cfU = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfU.Text = "U"; $cfU.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGoldenrodYellow; $cfU.Style.Font.Color.Color = [System.Drawing.Color]::DarkGoldenrod
                $cfD = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfD.Text = "D"; $cfD.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightCoral; $cfD.Style.Font.Color.Color = [System.Drawing.Color]::DarkRed
            }
            $pkg.Save()
            $pkg.Dispose()
            $hasBaseSheet = $true
            Write-Host "[Excel] サマリーExcel書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        }

        $sw.Restart()
        Write-Host "[Excel] 詳細シート作成中..." -ForegroundColor Cyan
        $detailMatrix = Build-CrudDetailMatrix -CrudResults $CrudResults -TableDefinitions $ddlForJa
        Write-Host "[Excel] 詳細構築完了: $($detailMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        if ($detailMatrix.Count -gt 0) {
            $detailParams = @{
                Path          = $ExcelPath
                WorksheetName = $DetailSheetName
                AutoSize      = $true
                BoldTopRow    = $true
            }
            if ($hasBaseSheet) { $detailParams.Append = $true }
            $detailParams.PassThru = $true
            $pkg = $detailMatrix | Export-Excel @detailParams
            $ws = $pkg.Workbook.Worksheets[$DetailSheetName]
            $ws.Row(1).Style.WrapText = $true
            $ws.Row(1).Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
            # Python(openpyxl) と同様: スクロール領域の左上を E2（1行目と A〜D 列を固定）
            $ws.View.FreezePanes(2, 5)
            $lastRow = $ws.Dimension.End.Row
            $lastCol = $ws.Dimension.End.Column
            if ($lastCol -ge 5) {
                $dataRange = [OfficeOpenXml.ExcelAddress]::new(2, 5, $lastRow, $lastCol)
                $cfC = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfC.Text = "C"; $cfC.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGreen; $cfC.Style.Font.Color.Color = [System.Drawing.Color]::DarkGreen
                $cfR = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfR.Text = "R"; $cfR.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightBlue; $cfR.Style.Font.Color.Color = [System.Drawing.Color]::DarkBlue
                $cfU = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfU.Text = "U"; $cfU.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightGoldenrodYellow; $cfU.Style.Font.Color.Color = [System.Drawing.Color]::DarkGoldenrod
                $cfD = $ws.ConditionalFormatting.AddContainsText($dataRange)
                $cfD.Text = "D"; $cfD.Style.Fill.BackgroundColor.Color = [System.Drawing.Color]::LightCoral; $cfD.Style.Font.Color.Color = [System.Drawing.Color]::DarkRed
            }
            $pkg.Save()
            $pkg.Dispose()
            $hasBaseSheet = $true
            Write-Host "[Excel] 詳細Excel書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        }
    }

    if ($null -ne $TableDefinitions -and $TableDefinitions.Count -gt 0) {
        Write-Host "[Excel] テーブル定義シート作成中..." -ForegroundColor Cyan
        $tableDefData = Build-TableDefSheet -TableDefinitions $TableDefinitions
        if ($tableDefData.Count -gt 0) {
            $tableParams = @{
                Path          = $ExcelPath
                WorksheetName = "テーブル定義"
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
                TableName     = "TableDefinitions"
            }
            if ($hasBaseSheet) { $tableParams.Append = $true }
            $tableDefData | Export-Excel @tableParams
            $hasBaseSheet = $true
        }
    }

    if ($null -ne $IndexDefinitions -and $IndexDefinitions.Count -gt 0) {
        Write-Host "[Excel] インデックス定義シート作成中..." -ForegroundColor Cyan
        $indexDefData = Build-IndexDefSheet -IndexDefinitions $IndexDefinitions
        if ($indexDefData.Count -gt 0) {
            $indexParams = @{
                Path          = $ExcelPath
                WorksheetName = "インデックス定義"
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
                TableName     = "IndexDefinitions"
            }
            if ($hasBaseSheet) { $indexParams.Append = $true }
            $indexDefData | Export-Excel @indexParams
            $hasBaseSheet = $true
        }
    }

    if ($null -ne $UnusedColumns -and $UnusedColumns.Count -gt 0) {
        Write-Host "[Excel] 未使用カラムシート作成中..." -ForegroundColor Cyan
        $unusedData = Build-UnusedColumnsSheet -UnusedColumns $UnusedColumns
        if ($unusedData.Count -gt 0) {
            $unusedParams = @{
                Path          = $ExcelPath
                WorksheetName = "未使用カラム"
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
                TableName     = "UnusedColumns"
            }
            if ($hasBaseSheet) { $unusedParams.Append = $true }
            $unusedData | Export-Excel @unusedParams
            $hasBaseSheet = $true
        }
    }

    if ($CrudResults.Count -gt 0) {
        $sw.Restart()
        Write-Host "[Excel] 生データシート作成中... ($($CrudResults.Count) 件)" -ForegroundColor Cyan
        $rawData = Build-RawDataSheet -CrudResults $CrudResults
        if ($rawData.Count -gt 0) {
            $rawParams = @{
                Path          = $ExcelPath
                WorksheetName = $RawSheetName
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
                TableName     = "CrudRawData"
            }
            if ($hasBaseSheet) { $rawParams.Append = $true }
            $rawData | Export-Excel @rawParams
            $hasBaseSheet = $true
            Write-Host "[Excel] 生データ書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        }
    }

    if (-not $hasBaseSheet) {
        Write-Warning "[Excel] 出力対象データがないため、Excelは作成されません。"
        return
    }

    Write-Host "[Excel] 出力完了: $ExcelPath" -ForegroundColor Green
}

function Export-CrudExcelWithCOM {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [string]$ExcelPath,
        [string]$SummarySheetName = "テーブル×機能サマリー",
        [string]$DetailSheetName = "項目別詳細",
        [string]$RawSheetName = "生データ",
        [System.Collections.ArrayList]$TableDefinitions = $null,
        [System.Collections.ArrayList]$IndexDefinitions = $null,
        [System.Collections.ArrayList]$UnusedColumns = $null
    )

    $ddlForJa = $null
    if ($null -ne $TableDefinitions -and $TableDefinitions.Count -gt 0) { $ddlForJa = $TableDefinitions }

    $outputDir = [System.IO.Path]::GetDirectoryName($ExcelPath)
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $excel = $null
    $workbook = $null

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Add()
        $baseSheet = $workbook.Worksheets.Item(1)
        $lastSheet = $baseSheet
        $hasSheet = $false

        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        if ($CrudResults.Count -gt 0) {
            Write-Host "[Excel/COM] サマリーシート作成中... ($($CrudResults.Count) 件)" -ForegroundColor Cyan
            $summaryMatrix = Build-CrudSummaryMatrix -CrudResults $CrudResults -TableDefinitions $ddlForJa
            Write-Host "[Excel/COM] サマリー構築完了: $($summaryMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            if ($summaryMatrix.Count -gt 0) {
                $baseSheet.Name = $SummarySheetName
                Write-MatrixToSheet -Sheet $baseSheet -Data $summaryMatrix -ApplyCrudColoring -CrudColorColumnStartIndex 2 -FreezePaneRow 2 -FreezePaneColumn 3
                $lastSheet = $baseSheet
                $hasSheet = $true
                Write-Host "[Excel/COM] サマリーシート書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            }

            $sw.Restart()
            Write-Host "[Excel/COM] 詳細シート作成中..." -ForegroundColor Cyan
            $detailMatrix = Build-CrudDetailMatrix -CrudResults $CrudResults -TableDefinitions $ddlForJa
            Write-Host "[Excel/COM] 詳細構築完了: $($detailMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            if ($detailMatrix.Count -gt 0) {
                $detailSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                $detailSheet.Name = $DetailSheetName
                Write-MatrixToSheet -Sheet $detailSheet -Data $detailMatrix -ApplyCrudColoring -CrudColorColumnStartIndex 4 -FreezePaneRow 2 -FreezePaneColumn 5
                $lastSheet = $detailSheet
                $hasSheet = $true
                Write-Host "[Excel/COM] 詳細シート書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            }
        }

        if ($null -ne $TableDefinitions -and $TableDefinitions.Count -gt 0) {
            Write-Host "[Excel/COM] テーブル定義シート作成中..." -ForegroundColor Cyan
            $tableDefData = Build-TableDefSheet -TableDefinitions $TableDefinitions
            if ($tableDefData.Count -gt 0) {
                $tableDefSheet = if ($hasSheet) {
                    $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                } else {
                    $baseSheet
                }
                $tableDefSheet.Name = "テーブル定義"
                Write-MatrixToSheet -Sheet $tableDefSheet -Data $tableDefData
                $lastSheet = $tableDefSheet
                $hasSheet = $true
            }
        }

        if ($null -ne $IndexDefinitions -and $IndexDefinitions.Count -gt 0) {
            Write-Host "[Excel/COM] インデックス定義シート作成中..." -ForegroundColor Cyan
            $indexDefData = Build-IndexDefSheet -IndexDefinitions $IndexDefinitions
            if ($indexDefData.Count -gt 0) {
                $indexDefSheet = if ($hasSheet) {
                    $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                } else {
                    $baseSheet
                }
                $indexDefSheet.Name = "インデックス定義"
                Write-MatrixToSheet -Sheet $indexDefSheet -Data $indexDefData
                $lastSheet = $indexDefSheet
                $hasSheet = $true
            }
        }

        if ($null -ne $UnusedColumns -and $UnusedColumns.Count -gt 0) {
            Write-Host "[Excel/COM] 未使用カラムシート作成中..." -ForegroundColor Cyan
            $unusedData = Build-UnusedColumnsSheet -UnusedColumns $UnusedColumns
            if ($unusedData.Count -gt 0) {
                $unusedSheet = if ($hasSheet) {
                    $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                } else {
                    $baseSheet
                }
                $unusedSheet.Name = "未使用カラム"
                Write-MatrixToSheet -Sheet $unusedSheet -Data $unusedData
                $lastSheet = $unusedSheet
                $hasSheet = $true
            }
        }

        if ($CrudResults.Count -gt 0) {
            $sw.Restart()
            Write-Host "[Excel/COM] 生データシート作成中... ($($CrudResults.Count) 件)" -ForegroundColor Cyan
            $rawData = Build-RawDataSheet -CrudResults $CrudResults
            if ($rawData.Count -gt 0) {
                $rawSheet = if ($hasSheet) {
                    $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                } else {
                    $baseSheet
                }
                $rawSheet.Name = $RawSheetName
                Write-MatrixToSheet -Sheet $rawSheet -Data $rawData
                $lastSheet = $rawSheet
                $hasSheet = $true
                Write-Host "[Excel/COM] 生データ書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            }
        }

        if (-not $hasSheet) {
            Write-Warning "[Excel/COM] 出力対象データがないため、Excelは作成されません。"
            return
        }

        $fullPath = [System.IO.Path]::GetFullPath($ExcelPath)
        $workbook.SaveAs($fullPath)
        Write-Host "[Excel/COM] 出力完了: $fullPath" -ForegroundColor Green
    }
    finally {
        if ($null -ne $workbook) {
            $workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($null -ne $excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

function Write-MatrixToSheet {
    param(
        $Sheet,
        [System.Collections.ArrayList]$Data,
        [switch]$ApplyCrudColoring,
        [int]$CrudColorColumnStartIndex = 1,
        [int]$FreezePaneRow = 2,
        [int]$FreezePaneColumn = 1
    )

    if ($Data.Count -eq 0) { return }

    $properties = $Data[0].PSObject.Properties.Name
    $rowCount = $Data.Count + 1
    $colCount = @($properties).Count

    Write-Host "  2D配列構築中 ($($Data.Count) 行 x $colCount 列)..." -ForegroundColor Gray
    $values = New-Object 'object[,]' $rowCount, $colCount

    $c = 0
    foreach ($prop in $properties) {
        $values[0, $c] = $prop
        $c++
    }

    $r = 1
    foreach ($item in $Data) {
        $c = 0
        foreach ($prop in $properties) {
            $values[$r, $c] = $item.$prop
            $c++
        }
        if ($r % 5000 -eq 0) {
            Write-Host "  配列構築中: $r / $($Data.Count)" -ForegroundColor Gray
        }
        $r++
    }

    Write-Host "  シートへ一括書込中..." -ForegroundColor Gray
    $range = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item($rowCount, $colCount))
    $range.Value2 = $values

    $headerRange = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item(1, $colCount))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.ColorIndex = 15
    $headerRange.WrapText = $true
    $headerRange.VerticalAlignment = -4160

    if ($ApplyCrudColoring -and $colCount -gt $CrudColorColumnStartIndex) {
        Write-Host "  CRUD色付け中..." -ForegroundColor Gray
        for ($r = 1; $r -lt $rowCount; $r++) {
            for ($c = $CrudColorColumnStartIndex; $c -lt $colCount; $c++) {
                $val = $values[$r, $c]
                if ($null -ne $val -and $val -ne '-') {
                    if ($val -match 'C') {
                        $Sheet.Cells.Item($r + 1, $c + 1).Interior.Color = 0xCCFFCC
                    }
                    if ($val -match 'D') {
                        $Sheet.Cells.Item($r + 1, $c + 1).Interior.Color = 0xCCCCFF
                    }
                }
            }
            if ($r % 1000 -eq 0) {
                Write-Host "  色付け中: $r / $($rowCount - 1) 行" -ForegroundColor Gray
            }
        }
    }

    $Sheet.UsedRange.Columns.AutoFit() | Out-Null

    # Python(openpyxl) の freeze_panes と同じ考え方: 指定セルがスクロール領域の左上（例: C2=行2列3、E2=行2列5）
    $Sheet.Activate() | Out-Null
    $Sheet.Application.ActiveWindow.FreezePanes = $false
    $Sheet.Cells.Item($FreezePaneRow, $FreezePaneColumn).Select() | Out-Null
    $Sheet.Application.ActiveWindow.FreezePanes = $true
}

function Export-CrudJson {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [string]$JsonPath
    )

    $outputDir = [System.IO.Path]::GetDirectoryName($JsonPath)
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $CrudResults | ConvertTo-Json -Depth 5 | Out-File -FilePath $JsonPath -Encoding UTF8
    Write-Host "[JSON] 出力完了: $JsonPath" -ForegroundColor Green
}
