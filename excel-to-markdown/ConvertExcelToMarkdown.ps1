<#
.SYNOPSIS
    Excel設計書（バッチ処理）をMarkdownに変換するプロトタイプスクリプト

.DESCRIPTION
    お試し変換用。設定セクションのパラメータを調整しながら繰り返し実行する。
    - 取り消し線付きテキストを判別して分離
    - 表形式のシートをMarkdownテーブルに変換
    - 非表示シートをスキップ
    - 結合セルを考慮

.USAGE
    .\ConvertExcelToMarkdown.ps1
    または設定セクションの $InputFile / $OutputFile を変更して実行
#>

# ============================================================
# 設定セクション（★ここを調整する）
# ============================================================

# 入力Excelファイルのパス
$InputFile = "C:\path\to\設計書.xlsx"

# 出力Markdownファイルのパス
$OutputFile = "C:\path\to\output.md"

# 表紙シートの設定（機能名を取得するセル位置）
$CoverSheet = @{
    Name       = "表紙"         # シート名（部分一致で検索）
    FuncNameCell = @(2, 2)      # 機能名が入っているセル [行, 列] （例：B2）
}

# 変換対象外のシート名パターン（部分一致）
$SkipSheetPatterns = @("変更履歴", "改版履歴")

# 表形式として扱うシート名パターン（部分一致）
$TableSheetPatterns = @("項目編集")

# 取り消し線テキストの扱い: "exclude"=除外 / "collapse"=折りたたみに退避
$StrikethroughMode = "collapse"

# デバッグモード: $true にするとセル読み取り時の詳細をコンソールに出力
$DebugMode = $true

# ============================================================
# 関数定義
# ============================================================

function Get-CellTextWithStrikethrough {
    <#
    .SYNOPSIS
        セルからテキストを取得し、取り消し線あり/なしで分離して返す
    #>
    param(
        [Parameter(Mandatory)]$Cell
    )

    $result = @{
        ActiveText = ""
        StrikeText = ""
    }

    if ($null -eq $Cell.Value2) {
        return $result
    }

    $totalLen = 0
    try {
        $totalLen = $Cell.Characters().Count
    } catch {
        $result.ActiveText = [string]$Cell.Value2
        return $result
    }

    if ($totalLen -eq 0) {
        return $result
    }

    # セル全体の取り消し線を先にチェック（高速化）
    $wholeStrike = $Cell.Font.Strikethrough
    if ($wholeStrike -eq $true) {
        $result.StrikeText = [string]$Cell.Value2
        return $result
    }
    if ($wholeStrike -eq $false) {
        $result.ActiveText = [string]$Cell.Value2
        return $result
    }

    # 部分的に取り消し線が混在（$wholeStrike が $null の場合）
    $activeChars = New-Object System.Text.StringBuilder
    $strikeChars = New-Object System.Text.StringBuilder

    for ($i = 1; $i -le $totalLen; $i++) {
        $charObj = $Cell.Characters($i, 1)
        $charText = $charObj.Text
        if ($charObj.Font.Strikethrough -eq $true) {
            [void]$strikeChars.Append($charText)
        } else {
            [void]$activeChars.Append($charText)
        }
    }

    $result.ActiveText = $activeChars.ToString()
    $result.StrikeText = $strikeChars.ToString()
    return $result
}

function Get-MergedCellValue {
    <#
    .SYNOPSIS
        結合セルを考慮してセルの値を取得する。
        結合範囲の左上セル以外は空文字を返す。
    #>
    param(
        [Parameter(Mandatory)]$Cell
    )

    if ($Cell.MergeCells) {
        $mergeArea = $Cell.MergeArea
        $topLeft = $mergeArea.Cells.Item(1, 1)
        if ($Cell.Row -eq $topLeft.Row -and $Cell.Column -eq $topLeft.Column) {
            return (Get-CellTextWithStrikethrough -Cell $Cell)
        } else {
            return @{ ActiveText = ""; StrikeText = "" }
        }
    }

    return (Get-CellTextWithStrikethrough -Cell $Cell)
}

function ConvertTo-MarkdownTable {
    <#
    .SYNOPSIS
        シートの使用範囲をMarkdownテーブルに変換する
    #>
    param(
        [Parameter(Mandatory)]$Sheet
    )

    $range = $Sheet.UsedRange
    if ($null -eq $range) { return "" }

    $rowCount = $range.Rows.Count
    $colCount = $range.Columns.Count
    $startRow = $range.Row
    $startCol = $range.Column

    if ($DebugMode) {
        Write-Host "  [TABLE] 範囲: $rowCount 行 x $colCount 列 (開始: R${startRow}C${startCol})" -ForegroundColor Cyan
    }

    $activeLines = @()
    $strikeTexts = @()
    $isFirstDataRow = $true

    for ($r = 1; $r -le $rowCount; $r++) {
        $rowCells = @()
        $rowHasContent = $false
        $rowStrikeTexts = @()

        for ($c = 1; $c -le $colCount; $c++) {
            $cell = $range.Cells.Item($r, $c)
            $textInfo = Get-MergedCellValue -Cell $cell

            $activeText = ($textInfo.ActiveText -replace "`r`n", " " -replace "`n", " ").Trim()
            $activeText = $activeText -replace "\|", "\|"

            if ($activeText -ne "") { $rowHasContent = $true }
            $rowCells += $activeText

            if ($textInfo.StrikeText -ne "") {
                $rowStrikeTexts += $textInfo.StrikeText.Trim()
            }
        }

        if (-not $rowHasContent) { continue }

        $line = "| " + ($rowCells -join " | ") + " |"
        $activeLines += $line

        if ($isFirstDataRow) {
            $separator = "|" + ("---|" * $colCount)
            $activeLines += $separator
            $isFirstDataRow = $false
        }

        if ($rowStrikeTexts.Count -gt 0) {
            foreach ($st in $rowStrikeTexts) {
                $strikeTexts += "- ~~${st}~~"
            }
        }
    }

    $md = ($activeLines -join "`n")

    if ($strikeTexts.Count -gt 0 -and $StrikethroughMode -eq "collapse") {
        $md += "`n`n<details>`n<summary>取り消し済みの記述</summary>`n`n"
        $md += ($strikeTexts -join "`n")
        $md += "`n`n</details>"
    }

    return $md
}

function ConvertTo-MarkdownText {
    <#
    .SYNOPSIS
        シートの使用範囲をMarkdownテキスト（文章形式）に変換する
    #>
    param(
        [Parameter(Mandatory)]$Sheet
    )

    $range = $Sheet.UsedRange
    if ($null -eq $range) { return "" }

    $rowCount = $range.Rows.Count
    $colCount = $range.Columns.Count

    if ($DebugMode) {
        Write-Host "  [TEXT] 範囲: $rowCount 行 x $colCount 列" -ForegroundColor Cyan
    }

    $activeLines = @()
    $strikeTexts = @()

    for ($r = 1; $r -le $rowCount; $r++) {
        $rowTexts = @()

        for ($c = 1; $c -le $colCount; $c++) {
            $cell = $range.Cells.Item($r, $c)
            $textInfo = Get-MergedCellValue -Cell $cell

            if ($textInfo.ActiveText -ne "") {
                $rowTexts += $textInfo.ActiveText.Trim()
            }
            if ($textInfo.StrikeText -ne "") {
                $strikeTexts += "- ~~" + $textInfo.StrikeText.Trim() + "~~"
            }
        }

        if ($rowTexts.Count -gt 0) {
            $lineText = $rowTexts -join "　"
            $activeLines += $lineText
        }
    }

    $md = ($activeLines -join "`n`n")

    if ($strikeTexts.Count -gt 0 -and $StrikethroughMode -eq "collapse") {
        $md += "`n`n<details>`n<summary>取り消し済みの記述</summary>`n`n"
        $md += ($strikeTexts -join "`n")
        $md += "`n`n</details>"
    }

    return $md
}

function Test-SheetNameMatch {
    <#
    .SYNOPSIS
        シート名がパターン一覧のいずれかに部分一致するか判定する
    #>
    param(
        [string]$SheetName,
        [string[]]$Patterns
    )
    foreach ($pattern in $Patterns) {
        if ($SheetName -like "*${pattern}*") { return $true }
    }
    return $false
}

# ============================================================
# メイン処理
# ============================================================

Write-Host "========================================" -ForegroundColor Green
Write-Host " Excel → Markdown 変換（お試し）" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "入力: $InputFile"
Write-Host "出力: $OutputFile"
Write-Host ""

# 入力ファイルの存在確認
if (-not (Test-Path $InputFile)) {
    Write-Host "[ERROR] ファイルが見つかりません: $InputFile" -ForegroundColor Red
    exit 1
}

$excel = $null
$workbook = $null

try {
    # Excel起動
    Write-Host "Excel を起動しています..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $fullPath = (Resolve-Path $InputFile).Path
    $workbook = $excel.Workbooks.Open($fullPath)

    Write-Host "シート数: $($workbook.Sheets.Count)" -ForegroundColor Yellow
    Write-Host ""

    # --- 表紙シートからメタ情報を取得 ---
    $funcName = "（不明）"

    foreach ($sheet in $workbook.Sheets) {
        if ($sheet.Name -like "*$($CoverSheet.Name)*") {
            $r = $CoverSheet.FuncNameCell[0]
            $c = $CoverSheet.FuncNameCell[1]
            $val = $sheet.Cells.Item($r, $c).Value2
            if ($val) { $funcName = [string]$val }
            Write-Host "[表紙] 機能名: $funcName" -ForegroundColor Green
            break
        }
    }

    # --- Markdown組み立て ---
    $mdContent = @()

    # YAMLフロントマター
    $mdContent += "---"
    $mdContent += "種別: バッチ処理設計書"
    $mdContent += "機能名: $funcName"
    $mdContent += "実装: PL/SQL"
    $mdContent += "元ファイル: $($InputFile | Split-Path -Leaf)"
    $mdContent += "変換日: $(Get-Date -Format 'yyyy-MM-dd')"
    $mdContent += "---"
    $mdContent += ""
    $mdContent += "# $funcName"
    $mdContent += ""

    # --- 各シートを処理 ---
    foreach ($sheet in $workbook.Sheets) {
        $sheetName = $sheet.Name

        # 非表示シートをスキップ（xlSheetVisible = -1）
        if ($sheet.Visible -ne -1) {
            Write-Host "[$sheetName] スキップ（非表示）" -ForegroundColor DarkGray
            continue
        }

        # 表紙シートはフロントマターで処理済み
        if ($sheetName -like "*$($CoverSheet.Name)*") {
            Write-Host "[$sheetName] スキップ（表紙→フロントマターに反映済み）" -ForegroundColor DarkGray
            continue
        }

        # 変換対象外シートをスキップ
        if (Test-SheetNameMatch -SheetName $sheetName -Patterns $SkipSheetPatterns) {
            Write-Host "[$sheetName] スキップ（対象外パターン）" -ForegroundColor DarkGray
            continue
        }

        Write-Host "[$sheetName] 変換中..." -ForegroundColor Yellow

        $mdContent += "## $sheetName"
        $mdContent += ""

        # 表形式シートか文章形式シートかを判定して変換
        if (Test-SheetNameMatch -SheetName $sheetName -Patterns $TableSheetPatterns) {
            $tableMd = ConvertTo-MarkdownTable -Sheet $sheet
            $mdContent += $tableMd
        } else {
            $textMd = ConvertTo-MarkdownText -Sheet $sheet
            $mdContent += $textMd
        }

        $mdContent += ""
        $mdContent += "---"
        $mdContent += ""
    }

    # --- ファイル出力 ---
    $outputText = $mdContent -join "`n"
    $outputText | Out-File -FilePath $OutputFile -Encoding UTF8

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " 変換完了: $OutputFile" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green

} catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
} finally {
    # Excel後処理
    if ($workbook) {
        $workbook.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Host "Excel を終了しました。" -ForegroundColor Yellow
}
