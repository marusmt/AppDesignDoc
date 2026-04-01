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

function Build-CrudSummaryMatrix {
    param([System.Collections.ArrayList]$CrudResults)

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

    $features = $featureSet | Sort-Object
    $tables = $tableSet | Sort-Object

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

        $row = [ordered]@{ "テーブル名" = $table }
        foreach ($feature in $features) {
            $header = $headerMap[$feature]
            $key = "$table|$feature"
            if ($lookup.ContainsKey($key)) {
                $row[$header] = ($lookup[$key] | Sort-Object) -join ""
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
    param([System.Collections.ArrayList]$CrudResults)

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

    $features = $featureSet | Sort-Object
    $tableColumnPairs = $pairSet | Sort-Object

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

        $row = [ordered]@{
            "テーブル名" = $table
            "項目名"     = $column
        }

        foreach ($feature in $features) {
            $header = $headerMap[$feature]
            $key = "$pair|$feature"
            if ($lookup.ContainsKey($key)) {
                $row[$header] = ($lookup[$key] | Sort-Object) -join ""
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
    foreach ($def in ($TableDefinitions | Sort-Object { $_.TableName }, { $_.OrdinalPos })) {
        [void]$rows.Add([PSCustomObject][ordered]@{
            "テーブル名" = $def.TableName
            "No"         = $def.OrdinalPos
            "カラム名"   = $def.ColumnName
            "データ型"   = $def.DataType
            "NULL許可"   = $def.Nullable
            "DEFAULT"    = $def.HasDefault
            "ソースファイル" = $def.SourceFile
        })
    }
    return $rows
}

function Build-IndexDefSheet {
    param([System.Collections.ArrayList]$IndexDefinitions)

    $rows = [System.Collections.ArrayList]::new()
    foreach ($def in ($IndexDefinitions | Sort-Object { $_.TableName }, { $_.IndexName }, { $_.ColumnPos })) {
        [void]$rows.Add([PSCustomObject][ordered]@{
            "テーブル名"     = $def.TableName
            "インデックス名" = $def.IndexName
            "一意性"         = $def.Uniqueness
            "カラム位置"     = $def.ColumnPos
            "カラム名"       = $def.ColumnName
            "ソースファイル" = $def.SourceFile
        })
    }
    return $rows
}

function Build-UnusedColumnsSheet {
    param([System.Collections.ArrayList]$UnusedColumns)

    $rows = [System.Collections.ArrayList]::new()
    foreach ($item in ($UnusedColumns | Sort-Object { $_.TableName }, { $_.ColumnName })) {
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
        $summaryMatrix = Build-CrudSummaryMatrix -CrudResults $CrudResults
        Write-Host "[Excel] サマリー構築完了: $($summaryMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        if ($summaryMatrix.Count -gt 0) {
            $pkg = $summaryMatrix | Export-Excel -Path $ExcelPath -WorksheetName $SummarySheetName `
                -AutoSize -FreezeTopRow -BoldTopRow -PassThru
            $ws = $pkg.Workbook.Worksheets[$SummarySheetName]
            $ws.Row(1).Style.WrapText = $true
            $ws.Row(1).Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
            $lastRow = $ws.Dimension.End.Row
            $lastCol = $ws.Dimension.End.Column
            if ($lastCol -ge 2) {
                $dataRange = [OfficeOpenXml.ExcelAddress]::new(2, 2, $lastRow, $lastCol)
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
        $detailMatrix = Build-CrudDetailMatrix -CrudResults $CrudResults
        Write-Host "[Excel] 詳細構築完了: $($detailMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
        if ($detailMatrix.Count -gt 0) {
            $detailParams = @{
                Path          = $ExcelPath
                WorksheetName = $DetailSheetName
                AutoSize      = $true
                FreezeTopRow  = $true
                BoldTopRow    = $true
            }
            if ($hasBaseSheet) { $detailParams.Append = $true }
            $detailParams.PassThru = $true
            $pkg = $detailMatrix | Export-Excel @detailParams
            $ws = $pkg.Workbook.Worksheets[$DetailSheetName]
            $ws.Row(1).Style.WrapText = $true
            $ws.Row(1).Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
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
            $summaryMatrix = Build-CrudSummaryMatrix -CrudResults $CrudResults
            Write-Host "[Excel/COM] サマリー構築完了: $($summaryMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            if ($summaryMatrix.Count -gt 0) {
                $baseSheet.Name = $SummarySheetName
                Write-MatrixToSheet -Sheet $baseSheet -Data $summaryMatrix -ApplyCrudColoring
                $lastSheet = $baseSheet
                $hasSheet = $true
                Write-Host "[Excel/COM] サマリーシート書込完了 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            }

            $sw.Restart()
            Write-Host "[Excel/COM] 詳細シート作成中..." -ForegroundColor Cyan
            $detailMatrix = Build-CrudDetailMatrix -CrudResults $CrudResults
            Write-Host "[Excel/COM] 詳細構築完了: $($detailMatrix.Count) 行 ($([int]$sw.Elapsed.TotalSeconds) 秒)" -ForegroundColor Cyan
            if ($detailMatrix.Count -gt 0) {
                $detailSheet = $workbook.Worksheets.Add([System.Reflection.Missing]::Value, $lastSheet)
                $detailSheet.Name = $DetailSheetName
                Write-MatrixToSheet -Sheet $detailSheet -Data $detailMatrix -ApplyCrudColoring
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
        [switch]$ApplyCrudColoring
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

    if ($ApplyCrudColoring -and $colCount -gt 1) {
        Write-Host "  CRUD色付け中..." -ForegroundColor Gray
        for ($r = 1; $r -lt $rowCount; $r++) {
            for ($c = 1; $c -lt $colCount; $c++) {
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

    $Sheet.Application.ActiveWindow.FreezePanes = $false
    $Sheet.Cells.Item(2, 1).Select() | Out-Null
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
