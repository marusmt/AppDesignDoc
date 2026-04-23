<#
.SYNOPSIS
    Oracle テーブル定義（DDL）・インデックス定義を解析する

.DESCRIPTION
    CREATE TABLE / CREATE INDEX 文を含む .sql ファイルを読み込み、
    テーブル名・カラム名・データ型・制約・インデックス情報を抽出する
    解析結果は SELECT * 展開やCRUDマトリックスの補完に利用する
#>

function Read-SqlFileAutoEncoding {
    param(
        [string]$Path,
        [string]$Encoding = "auto"
    )
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $encLower = if ([string]::IsNullOrWhiteSpace($Encoding)) { "auto" } else { $Encoding.Trim().ToLower() }
    if ($encLower -ne "auto") {
        $encName = if ($encLower -in @("shift_jis", "shift-jis", "sjis", "shiftjis")) { "shift_jis" } else { $encLower }
        return [System.Text.Encoding]::GetEncoding($encName).GetString($bytes)
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    try {
        $utf8Strict = New-Object System.Text.UTF8Encoding($false, $true)
        return $utf8Strict.GetString($bytes)
    }
    catch {
        return [System.Text.Encoding]::GetEncoding('shift_jis').GetString($bytes)
    }
}

function Remove-SqlCommentsForDdl {
    param([string]$Content)

    $result = $Content -replace '--[^\r\n]*', ''
    $result = $result -replace '/\*[\s\S]*?\*/', ''
    return $result
}

function Unescape-OracleDdlCommentString {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    return $Text -replace "''", "'"
}

function Parse-OracleCommentStatements {
    param([string]$Content)

    $tableComments = @{}
    $columnComments = @{}

    $patternTable = '(?is)COMMENT\s+ON\s+TABLE\s+(?:(\w+)\.)?(\w+)\s+IS\s*''((?:[^'']|'')*?)''\s*;'
    foreach ($m in [regex]::Matches($Content, $patternTable)) {
        $schema = if ($m.Groups[1].Success) { $m.Groups[1].Value.ToUpper() } else { "" }
        $tbl = $m.Groups[2].Value.ToUpper()
        $txt = Unescape-OracleDdlCommentString $m.Groups[3].Value
        $key = if ($schema -ne '') { "$schema.$tbl" } else { $tbl }
        $tableComments[$key] = $txt
    }

    $patternCol = '(?is)COMMENT\s+ON\s+COLUMN\s+(?:(\w+)\.)?(\w+)\.(\w+)\s+IS\s*''((?:[^'']|'')*?)''\s*;'
    foreach ($m in [regex]::Matches($Content, $patternCol)) {
        $schema = if ($m.Groups[1].Success) { $m.Groups[1].Value.ToUpper() } else { "" }
        $tbl = $m.Groups[2].Value.ToUpper()
        $col = $m.Groups[3].Value.ToUpper()
        $txt = Unescape-OracleDdlCommentString $m.Groups[4].Value
        $key = if ($schema -ne '') { "$schema.$tbl|$col" } else { "$tbl|$col" }
        $columnComments[$key] = $txt
    }

    return @{
        TableComments  = $tableComments
        ColumnComments = $columnComments
    }
}

function Merge-DdlCommentsIntoTableDefinitions {
    param(
        [System.Collections.ArrayList]$TableDefinitions,
        [hashtable]$TableComments,
        [hashtable]$ColumnComments
    )

    foreach ($def in $TableDefinitions) {
        $schema = if ($null -ne $def.Schema -and $def.Schema -ne '') { $def.Schema } else { "" }
        $tbl = $def.TableName
        $tblKey = if ($schema -ne '') { "$schema.$tbl" } else { $tbl }
        $def.TableComment = if ($TableComments.ContainsKey($tblKey)) { $TableComments[$tblKey] }
        elseif ($TableComments.ContainsKey($tbl)) { $TableComments[$tbl] } else { "" }

        $ckFull = if ($schema -ne '') { "$schema.$tbl|$($def.ColumnName)" } else { "$tbl|$($def.ColumnName)" }
        $ckShort = "$tbl|$($def.ColumnName)"
        $def.ColumnComment = if ($ColumnComments.ContainsKey($ckFull)) { $ColumnComments[$ckFull] }
        elseif ($ColumnComments.ContainsKey($ckShort)) { $ColumnComments[$ckShort] } else { "" }
    }
}

function Merge-DdlCommentsIntoIndexDefinitions {
    param(
        [System.Collections.ArrayList]$IndexDefinitions,
        [hashtable]$TableComments,
        [hashtable]$ColumnComments
    )

    foreach ($def in $IndexDefinitions) {
        $ts = if ($null -ne $def.TableSchema -and $def.TableSchema -ne '') { $def.TableSchema } else { "" }
        $tbl = $def.TableName
        $col = $def.ColumnName
        $tblKey = if ($ts -ne '') { "$ts.$tbl" } else { $tbl }
        $def.TableComment = if ($TableComments.ContainsKey($tblKey)) { $TableComments[$tblKey] }
        elseif ($TableComments.ContainsKey($tbl)) { $TableComments[$tbl] } else { "" }

        $ckFull = if ($ts -ne '') { "$ts.$tbl|$col" } else { "$tbl|$col" }
        $ckShort = "$tbl|$col"
        $def.ColumnComment = if ($ColumnComments.ContainsKey($ckFull)) { $ColumnComments[$ckFull] }
        elseif ($ColumnComments.ContainsKey($ckShort)) { $ColumnComments[$ckShort] } else { "" }
    }
}

function Parse-CreateTable {
    param([string]$Content)

    $results = [System.Collections.ArrayList]::new()
    $cleaned = Remove-SqlCommentsForDdl -Content $Content

    $headerPattern = '(?i)CREATE\s+(?:GLOBAL\s+TEMPORARY\s+)?TABLE\s+"?(?:(\w+)"?\."?)?(\w+)"?\s*\('
    $headerMatches = [regex]::Matches($cleaned, $headerPattern)

    foreach ($match in $headerMatches) {
        $schema = if ($match.Groups[1].Success) { $match.Groups[1].Value.ToUpper() } else { "" }
        $tableName = $match.Groups[2].Value.ToUpper()

        $startPos = $match.Index + $match.Length
        $depth = 1
        $pos = $startPos
        $found = $false

        while ($pos -lt $cleaned.Length) {
            $ch = $cleaned[$pos]
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') {
                $depth--
                if ($depth -eq 0) {
                    $found = $true
                    break
                }
            }
            $pos++
        }

        if (-not $found) { continue }

        $body = $cleaned.Substring($startPos, $pos - $startPos)
        $columns = Extract-ColumnDefinitions -TableBody $body -TableName $tableName -Schema $schema

        foreach ($col in $columns) {
            [void]$results.Add($col)
        }
    }

    return $results
}

function Extract-ColumnDefinitions {
    param([string]$TableBody, [string]$TableName, [string]$Schema)

    $columns = [System.Collections.ArrayList]::new()

    $depth = 0
    $current = ""
    $parts = [System.Collections.ArrayList]::new()

    foreach ($char in $TableBody.ToCharArray()) {
        if ($char -eq '(') { $depth++ }
        elseif ($char -eq ')') { $depth-- }

        if ($char -eq ',' -and $depth -eq 0) {
            [void]$parts.Add($current.Trim())
            $current = ""
        }
        else {
            $current += $char
        }
    }
    if ($current.Trim() -ne '') { [void]$parts.Add($current.Trim()) }

    $ordinalPos = 0
    foreach ($part in $parts) {
        $trimmed = $part.Trim()

        if ($trimmed -match '(?i)^\s*CONSTRAINT\b') { continue }
        if ($trimmed -match '(?i)^\s*PRIMARY\s+KEY\b') { continue }
        if ($trimmed -match '(?i)^\s*FOREIGN\s+KEY\b') { continue }
        if ($trimmed -match '(?i)^\s*UNIQUE\b') { continue }
        if ($trimmed -match '(?i)^\s*CHECK\b') { continue }
        if ($trimmed -match '(?i)^\s*SUPPLEMENTAL\b') { continue }

        if ($trimmed -match '(?i)^\s*"?(\w+)"?\s+(VARCHAR2|NVARCHAR2|CHAR|NCHAR|NUMBER|INTEGER|FLOAT|DATE|TIMESTAMP|CLOB|NCLOB|BLOB|RAW|LONG|XMLTYPE|ROWID|BINARY_FLOAT|BINARY_DOUBLE|INTERVAL)') {
            $ordinalPos++
            $colName = $Matches[1].ToUpper()
            $dataType = $Matches[2].ToUpper()

            $fullType = $dataType
            if ($trimmed -match "(?i)$dataType\s*\(([^)]+)\)") {
                $fullType = "$dataType($($Matches[1]))"
            }

            $nullable = if ($trimmed -match '(?i)\bNOT\s+NULL\b') { "NOT NULL" } else { "NULL" }
            $hasDefault = if ($trimmed -match '(?i)\bDEFAULT\b') { "YES" } else { "NO" }

            [void]$columns.Add(@{
                Schema        = $Schema
                TableName     = $TableName
                ColumnName    = $colName
                DataType      = $fullType
                Nullable      = $nullable
                HasDefault    = $hasDefault
                OrdinalPos    = $ordinalPos
                TableComment  = ""
                ColumnComment = ""
            })
        }
    }

    return $columns
}

function Extract-PrimaryKeyDefinitionsFromTableBody {
    param(
        [string]$TableBody,
        [string]$TableName,
        [string]$Schema
    )

    $results = [System.Collections.ArrayList]::new()
    $m = [regex]::Match($TableBody, '(?is)CONSTRAINT\s+(\w+)\s+PRIMARY\s+KEY\s*\(([^)]+)\)')
    if ($m.Success) {
        $pkName = $m.Groups[1].Value.ToUpper()
        $colsRaw = $m.Groups[2].Value
    }
    else {
        $m = [regex]::Match($TableBody, '(?is)\bPRIMARY\s+KEY\s*\(([^)]+)\)')
        if (-not $m.Success) { return $results }
        $pkName = "PK_$TableName"
        $colsRaw = $m.Groups[1].Value
    }

    $colParts = $colsRaw -split ','
    $colPos = 0
    foreach ($part in $colParts) {
        $col = $part.Trim() -replace '"', ''
        $col = $col.ToUpper()
        $col = $col -replace '\s+(ASC|DESC)\s*$', ''
        if ($col -eq '') { continue }
        $colPos++
        [void]$results.Add(@{
            IndexSchema    = ""
            IndexName      = $pkName
            TableSchema    = $Schema
            TableName      = $TableName
            ColumnName     = $col
            ColumnPos      = $colPos
            Uniqueness     = "UNIQUE"
            DefinitionKind = "PK"
            TableComment   = ""
            ColumnComment  = ""
        })
    }

    return $results
}

function Parse-PrimaryKeyConstraints {
    param([string]$Content)

    $results = [System.Collections.ArrayList]::new()
    $cleaned = Remove-SqlCommentsForDdl -Content $Content

    $headerPattern = '(?i)CREATE\s+(?:GLOBAL\s+TEMPORARY\s+)?TABLE\s+"?(?:(\w+)"?\."?)?(\w+)"?\s*\('
    $headerMatches = [regex]::Matches($cleaned, $headerPattern)

    foreach ($match in $headerMatches) {
        $schema = if ($match.Groups[1].Success) { $match.Groups[1].Value.ToUpper() } else { "" }
        $tableName = $match.Groups[2].Value.ToUpper()

        $startPos = $match.Index + $match.Length
        $depth = 1
        $pos = $startPos
        $found = $false

        while ($pos -lt $cleaned.Length) {
            $ch = $cleaned[$pos]
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') {
                $depth--
                if ($depth -eq 0) {
                    $found = $true
                    break
                }
            }
            $pos++
        }

        if (-not $found) { continue }

        $body = $cleaned.Substring($startPos, $pos - $startPos)
        $pkRows = Extract-PrimaryKeyDefinitionsFromTableBody -TableBody $body -TableName $tableName -Schema $schema
        foreach ($r in $pkRows) {
            [void]$results.Add($r)
        }
    }

    return $results
}

function Parse-CreateIndex {
    param([string]$Content)

    $results = [System.Collections.ArrayList]::new()
    $cleaned = Remove-SqlCommentsForDdl -Content $Content

    $headerPattern = '(?i)CREATE\s+(UNIQUE\s+)?INDEX\s+"?(?:(\w+)"?\."?)?(\w+)"?\s+ON\s+"?(?:(\w+)"?\."?)?(\w+)"?\s*\('
    $headerMatches = [regex]::Matches($cleaned, $headerPattern)

    foreach ($match in $headerMatches) {
        $isUnique = if ($match.Groups[1].Success) { "UNIQUE" } else { "NONUNIQUE" }
        $indexSchema = if ($match.Groups[2].Success) { $match.Groups[2].Value.ToUpper() } else { "" }
        $indexName = $match.Groups[3].Value.ToUpper()
        $tableSchema = if ($match.Groups[4].Success) { $match.Groups[4].Value.ToUpper() } else { "" }
        $tableName = $match.Groups[5].Value.ToUpper()

        $startPos = $match.Index + $match.Length
        $depth = 1
        $pos = $startPos
        $found = $false

        while ($pos -lt $cleaned.Length) {
            $ch = $cleaned[$pos]
            if ($ch -eq '(') { $depth++ }
            elseif ($ch -eq ')') {
                $depth--
                if ($depth -eq 0) {
                    $found = $true
                    break
                }
            }
            $pos++
        }

        if (-not $found) { continue }

        $columnsRaw = $cleaned.Substring($startPos, $pos - $startPos)

        $colDepth = 0
        $current = ""
        $columns = [System.Collections.ArrayList]::new()
        foreach ($ch in $columnsRaw.ToCharArray()) {
            if ($ch -eq '(') { $colDepth++ }
            elseif ($ch -eq ')') { $colDepth-- }

            if ($ch -eq ',' -and $colDepth -eq 0) {
                if ($current.Trim() -ne '') { [void]$columns.Add($current.Trim()) }
                $current = ""
            }
            else {
                $current += $ch
            }
        }
        if ($current.Trim() -ne '') { [void]$columns.Add($current.Trim()) }

        $colPos = 0
        foreach ($colExpr in $columns) {
            $colPos++
            $col = $colExpr.ToUpper()
            $col = $col -replace '\s+(ASC|DESC)\s*$', ''
            if ($col -match '^\w+\s*\((.+)\)$') {
                $col = $Matches[1].Trim() -replace '"', ''
            }
            $col = $col -replace '"', ''
            $col = $col.Trim()
            if ($col -eq '') { continue }
            [void]$results.Add(@{
                IndexSchema    = $indexSchema
                IndexName      = $indexName
                TableSchema    = $tableSchema
                TableName      = $tableName
                ColumnName     = $col
                ColumnPos      = $colPos
                Uniqueness     = $isUnique
                DefinitionKind = "INDEX"
                TableComment   = ""
                ColumnComment  = ""
            })
        }
    }

    return $results
}

function Parse-TableDdlDirectory {
    param(
        [string]$SourcePath,
        [string]$FilePattern = "*.sql",
        [string[]]$ExcludePatterns = @(),
        [string[]]$ExcludeTables = @(),
        [string]$SourceEncoding = "auto"
    )

    Write-Host "[DDL] テーブル定義解析開始: $SourcePath" -ForegroundColor Cyan

    $files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -Recurse -File
    foreach ($pattern in $ExcludePatterns) {
        $files = $files | Where-Object { $_.Name -notlike $pattern }
    }

    Write-Host "[DDL] 対象ファイル数: $($files.Count)" -ForegroundColor Cyan

    $allTableDefs = [System.Collections.ArrayList]::new()
    $allIndexDefs = [System.Collections.ArrayList]::new()
    $globalTableComments = @{}
    $globalColumnComments = @{}
    $fileCount = 0

    foreach ($file in $files) {
        $fileCount++
        Write-Progress -Activity "DDL 解析中" -Status "$fileCount / $($files.Count): $($file.Name)" -PercentComplete (($fileCount / $files.Count) * 100)

        try {
            $content = Read-SqlFileAutoEncoding -Path $file.FullName -Encoding $SourceEncoding

            $cstmt = Parse-OracleCommentStatements -Content $content
            foreach ($k in $cstmt.TableComments.Keys) { $globalTableComments[$k] = $cstmt.TableComments[$k] }
            foreach ($k in $cstmt.ColumnComments.Keys) { $globalColumnComments[$k] = $cstmt.ColumnComments[$k] }

            $tableDefs = Parse-CreateTable -Content $content
            foreach ($def in $tableDefs) {
                if ($def.TableName -notin $ExcludeTables) {
                    $def.SourceFile = $file.Name
                    [void]$allTableDefs.Add($def)
                }
            }

            $indexDefs = Parse-CreateIndex -Content $content
            foreach ($def in $indexDefs) {
                if ($def.TableName -notin $ExcludeTables) {
                    $def.SourceFile = $file.Name
                    [void]$allIndexDefs.Add($def)
                }
            }

            $pkDefs = Parse-PrimaryKeyConstraints -Content $content
            $pkSeen = @{}
            foreach ($def in $pkDefs) {
                if ($def.TableName -notin $ExcludeTables) {
                    $sig = "$($def.TableName)|$($def.IndexName)|$($def.ColumnPos)|$($def.ColumnName)"
                    if ($pkSeen.ContainsKey($sig)) { continue }
                    $pkSeen[$sig] = $true
                    $def.SourceFile = $file.Name
                    [void]$allIndexDefs.Add($def)
                }
            }
        }
        catch {
            Write-Warning "[DDL] 解析エラー: $($file.FullName) - $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "DDL 解析中" -Completed

    if ($allTableDefs.Count -gt 0) {
        Merge-DdlCommentsIntoTableDefinitions -TableDefinitions $allTableDefs -TableComments $globalTableComments -ColumnComments $globalColumnComments
    }
    if ($allIndexDefs.Count -gt 0) {
        Merge-DdlCommentsIntoIndexDefinitions -IndexDefinitions $allIndexDefs -TableComments $globalTableComments -ColumnComments $globalColumnComments
    }
    $pkDefCount = @($allIndexDefs | Where-Object { $_.DefinitionKind -eq 'PK' } | ForEach-Object { $_.IndexName } | Sort-Object -Unique).Count

    $tableCount = ($allTableDefs | ForEach-Object { $_.TableName } | Sort-Object -Unique).Count
    $columnCount = $allTableDefs.Count
    $indexCount = ($allIndexDefs | ForEach-Object { $_.IndexName } | Sort-Object -Unique).Count

    Write-Host "[DDL] 解析完了: テーブル $tableCount 件, カラム $columnCount 件, インデックス定義 $indexCount 件（主キー制約 $pkDefCount 件）" -ForegroundColor Green

    return @{
        TableDefinitions = $allTableDefs
        IndexDefinitions = $allIndexDefs
    }
}

function Expand-SelectStar {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [System.Collections.ArrayList]$TableDefinitions
    )

    $expanded = [System.Collections.ArrayList]::new()
    $expandedCount = 0

    foreach ($item in $CrudResults) {
        if ($item.ColumnName -eq "*" -and $item.Operation -eq "R") {
            $tableCols = $TableDefinitions | Where-Object { $_.TableName -eq $item.TableName }

            if ($tableCols.Count -gt 0) {
                $expandedCount++
                foreach ($col in $tableCols) {
                    $baseDtm = $true
                    if ($item -is [hashtable] -and $item.ContainsKey('DdlTableMatched')) {
                        $baseDtm = $item['DdlTableMatched']
                    }
                    [void]$expanded.Add(@{
                        SourceType  = $item.SourceType
                        SourceFile  = $item.SourceFile
                        ObjectType  = $item.ObjectType
                        ObjectName  = $item.ObjectName
                        ProcName    = $item.ProcName
                        FeatureName = $item.FeatureName
                        TableName   = $item.TableName
                        ColumnName  = $col.ColumnName
                        Operation   = "R"
                        DdlTableMatched = $baseDtm
                    })
                }
            }
            else {
                if ($item -is [hashtable] -and -not $item.ContainsKey('DdlTableMatched')) {
                    $item['DdlTableMatched'] = $true
                }
                [void]$expanded.Add($item)
            }
        }
        else {
            if ($item -is [hashtable] -and -not $item.ContainsKey('DdlTableMatched')) {
                $item['DdlTableMatched'] = $true
            }
            [void]$expanded.Add($item)
        }
    }

    if ($expandedCount -gt 0) {
        Write-Host "[DDL] SELECT * 展開: $expandedCount 箇所を個別カラムに展開" -ForegroundColor Green
    }

    return $expanded
}

function Test-ColumnExistence {
    param(
        [System.Collections.ArrayList]$CrudResults,
        [System.Collections.ArrayList]$TableDefinitions
    )

    $ddlColumns = @{}
    $ddlTables = @{}
    foreach ($def in $TableDefinitions) {
        $ddlTables[$def.TableName] = $true
        $key = "$($def.TableName)|$($def.ColumnName)"
        $ddlColumns[$key] = $true
    }

    $validated = [System.Collections.ArrayList]::new()
    $removed = [System.Collections.ArrayList]::new()
    $skipColumns = @('*', '(ALL)')

    foreach ($item in $CrudResults) {
        if ($item.ColumnName -in $skipColumns) {
            [void]$validated.Add($item)
            continue
        }

        if (-not $ddlTables.ContainsKey($item.TableName)) {
            [void]$validated.Add($item)
            continue
        }

        $key = "$($item.TableName)|$($item.ColumnName)"
        if ($ddlColumns.ContainsKey($key)) {
            [void]$validated.Add($item)
        }
        else {
            [void]$removed.Add($item)
        }
    }

    $removedCount = $removed.Count
    if ($removedCount -gt 0) {
        $removedTables = ($removed | ForEach-Object { $_.TableName } | Sort-Object -Unique).Count
        Write-Host "[検証] DDL突合せ: $removedCount 件の不正エントリを除外（$removedTables テーブル）" -ForegroundColor Yellow

        $removedGroups = $removed | Group-Object { "$($_.TableName)|$($_.ColumnName)" }
        $displayLimit = [Math]::Min($removedGroups.Count, 20)
        for ($i = 0; $i -lt $displayLimit; $i++) {
            $grp = $removedGroups[$i]
            $parts = $grp.Name -split '\|'
            Write-Host "        除外: $($parts[0]).$($parts[1]) ($($grp.Count) 件)" -ForegroundColor DarkGray
        }
        if ($removedGroups.Count -gt 20) {
            Write-Host "        ... 他 $($removedGroups.Count - 20) 件" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "[検証] DDL突合せ: 不正エントリなし" -ForegroundColor Green
    }

    return @{
        Validated = $validated
        Removed   = $removed
    }
}

function Find-UnusedColumns {
    param(
        [System.Collections.ArrayList]$TableDefinitions,
        [System.Collections.ArrayList]$CrudResults
    )

    $usedColumns = @{}
    foreach ($item in $CrudResults) {
        $key = "$($item.TableName)|$($item.ColumnName)"
        $usedColumns[$key] = $true
    }

    $unused = [System.Collections.ArrayList]::new()
    foreach ($def in $TableDefinitions) {
        $key = "$($def.TableName)|$($def.ColumnName)"
        if (-not $usedColumns.ContainsKey($key)) {
            [void]$unused.Add(@{
                TableName  = $def.TableName
                ColumnName = $def.ColumnName
                DataType   = $def.DataType
                Nullable   = $def.Nullable
            })
        }
    }

    $unusedTableCount = ($unused | ForEach-Object { $_.TableName } | Sort-Object -Unique).Count
    Write-Host "[分析] 未使用カラム: $($unused.Count) 件（$unusedTableCount テーブル）" -ForegroundColor Yellow

    return $unused
}
