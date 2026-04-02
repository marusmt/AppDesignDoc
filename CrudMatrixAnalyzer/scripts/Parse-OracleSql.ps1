<#
.SYNOPSIS
    Oracle SQL ソースファイルを解析し、CRUD操作情報を抽出する

.DESCRIPTION
    Oracle PL/SQL パッケージ・トリガー・ビュー等のSQLファイルを読み込み、
    INSERT/SELECT/UPDATE/DELETE/MERGE文からテーブル名・項目名・操作種別を抽出する

.PARAMETER SourcePath
    SQLファイル格納ディレクトリ

.PARAMETER FilePattern
    対象ファイルパターン（デフォルト: *.sql）

.PARAMETER ExcludePatterns
    除外ファイルパターン配列
#>

function Remove-SqlComments {
    param([string]$Content)

    $result = $Content -replace [char]0x3000, ' '
    $result = $result -replace '--.*?(?=\r?\n|$)', ''
    $result = $result -replace '/\*[\s\S]*?\*/', ''
    return $result
}

function Get-OracleObjectInfo {
    param([string]$Content, [string]$FileName)

    $objectName = ""
    $objectType = ""

    if ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+BODY\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "PACKAGE"
    }
    elseif ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+TRIGGER\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "TRIGGER"
    }
    elseif ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+(?:MATERIALIZED\s+)?VIEW\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "VIEW"
    }
    elseif ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+FUNCTION\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "FUNCTION"
    }
    elseif ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+PROCEDURE\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "PROCEDURE"
    }
    elseif ($Content -match '(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+(?:\w+\.)?(\w+)') {
        $objectName = $Matches[1].ToUpper()
        $objectType = "PACKAGE_SPEC"
    }
    else {
        $objectName = [System.IO.Path]::GetFileNameWithoutExtension($FileName).ToUpper()
        $objectType = "OTHER"
    }

    return @{
        ObjectName = $objectName
        ObjectType = $objectType
    }
}

function Get-ProcedureBlocks {
    param([string]$Content)

    $blocks = [System.Collections.ArrayList]::new()

    $linePattern = '(?im)^(\s*)(?:PROCEDURE|FUNCTION)\s+(\w+)'
    $regexMatches = [regex]::Matches($Content, $linePattern)

    if ($regexMatches.Count -eq 0) {
        $fallbackPattern = '(?i)(?:PROCEDURE|FUNCTION)\s+(\w+)'
        $regexMatches = [regex]::Matches($Content, $fallbackPattern)
        if ($regexMatches.Count -eq 0) {
            [void]$blocks.Add(@{
                Name    = "(MAIN)"
                Content = $Content
            })
            return $blocks
        }

        for ($i = 0; $i -lt $regexMatches.Count; $i++) {
            $startIdx = $regexMatches[$i].Index
            $endIdx = if ($i + 1 -lt $regexMatches.Count) { $regexMatches[$i + 1].Index } else { $Content.Length }
            $blockContent = $Content.Substring($startIdx, $endIdx - $startIdx)
            $procName = $regexMatches[$i].Groups[1].Value.ToUpper()

            [void]$blocks.Add(@{
                Name    = $procName
                Content = $blockContent
            })
        }
        return $blocks
    }

    for ($i = 0; $i -lt $regexMatches.Count; $i++) {
        $startIdx = $regexMatches[$i].Index
        $endIdx = if ($i + 1 -lt $regexMatches.Count) { $regexMatches[$i + 1].Index } else { $Content.Length }
        $blockContent = $Content.Substring($startIdx, $endIdx - $startIdx)
        $procName = $regexMatches[$i].Groups[2].Value.ToUpper()

        [void]$blocks.Add(@{
            Name    = $procName
            Content = $blockContent
        })
    }

    return $blocks
}

function Get-DeleteCrudRows {
    param(
        [string]$SqlFragment,
        [System.Collections.Generic.HashSet[string]]$CteNames
    )

    $out = [System.Collections.ArrayList]::new()
    $seenTableNames = [System.Collections.ArrayList]::new()

    $patterns = @(
        '(?i)DELETE\s+FROM\s+(?:([\w$]+)\.)?([\w$]+)',
        '(?i)DELETE\s+(?:([\w$]+)\.)?([\w$]+)\s+WHERE\b',
        '(?i)DELETE\s+(?:([\w$]+)\.)?([\w$]+)\s*;'
    )
    foreach ($pat in $patterns) {
        $m = [regex]::Matches($SqlFragment, $pat)
        foreach ($match in $m) {
            $tableName = $match.Groups[2].Value.ToUpper()
            if ($tableName -eq '') { continue }
            if ($CteNames.Count -gt 0 -and $CteNames.Contains($tableName)) { continue }
            if ($seenTableNames -contains $tableName) { continue }
            [void]$seenTableNames.Add($tableName)
            [void]$out.Add(@{
                TableName  = $tableName
                ColumnName = "(ALL)"
                Operation  = "D"
            })
        }
    }

    return $out
}

function Get-TableAndColumns {
    param([string]$SqlFragment, [string]$OperationType)

    $SqlFragment = $SqlFragment -replace [char]0x3000, ' '

    $results = [System.Collections.ArrayList]::new()

    $cteNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($SqlFragment -match '(?i)\bWITH\b') {
        $cteMatches = [regex]::Matches($SqlFragment, '(?i)([\w$]+)\s+AS\s*\(')
        foreach ($cm in $cteMatches) {
            [void]$cteNames.Add($cm.Groups[1].Value.ToUpper())
        }
    }

    switch ($OperationType) {
        "INSERT" {
            $pattern = '(?i)INSERT\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)\s*\(([^)]+)\)'
            $m = [regex]::Matches($SqlFragment, $pattern)
            foreach ($match in $m) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                $columnsRaw = $match.Groups[3].Value
                $columns = ($columnsRaw -split ',') | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' }
                foreach ($col in $columns) {
                    [void]$results.Add(@{
                        TableName  = $tableName
                        ColumnName = $col
                        Operation  = "C"
                    })
                }
            }
        }
        "SELECT" {
            $selectMatches = [regex]::Matches($SqlFragment, '(?i)\bSELECT\b')
            foreach ($sm in $selectMatches) {
                $depth = 0
                $selectStart = $sm.Index + $sm.Length
                $fromStart = -1
                $fromEnd = -1
                $scanPos = $selectStart
                $text = $SqlFragment

                while ($scanPos -lt $text.Length) {
                    $ch = $text[$scanPos]
                    if ($ch -eq '(') { $depth++ }
                    elseif ($ch -eq ')') {
                        $depth--
                        if ($depth -lt 0) { break }
                    }

                    if ($depth -eq 0 -and $scanPos -lt $text.Length) {
                        $tail = $text.Substring($scanPos)
                        if ($tail -match '(?i)^FROM\b') {
                            $fromStart = $scanPos
                            break
                        }
                    }
                    $scanPos++
                }

                if ($fromStart -lt 0) { continue }

                $selectClause = $text.Substring($selectStart, $fromStart - $selectStart).Trim()
                $afterFrom = $fromStart + 4
                if ($afterFrom -ge $text.Length) { continue }

                $depth = 0
                $fromBodyStart = $afterFrom
                $fromEnd = $text.Length
                $scanPos = $afterFrom

                $terminators = @('WHERE', 'ORDER', 'GROUP', 'HAVING', 'UNION', 'INTERSECT', 'MINUS', 'FETCH', 'FOR')

                while ($scanPos -lt $text.Length) {
                    $ch = $text[$scanPos]
                    if ($ch -eq '(') { $depth++ }
                    elseif ($ch -eq ')') {
                        $depth--
                        if ($depth -lt 0) {
                            $fromEnd = $scanPos
                            break
                        }
                    }
                    elseif ($ch -eq ';' -and $depth -eq 0) {
                        $fromEnd = $scanPos
                        break
                    }

                    if ($depth -eq 0) {
                        foreach ($term in $terminators) {
                            $termLen = $term.Length
                            if (($scanPos + $termLen) -le $text.Length) {
                                $candidate = $text.Substring($scanPos, $termLen)
                                if ($candidate -match "(?i)^$term$") {
                                    if ($scanPos -gt 0 -and $text[$scanPos - 1] -match '\s') {
                                        if (($scanPos + $termLen) -ge $text.Length -or $text[$scanPos + $termLen] -match '[\s(;]') {
                                            $fromEnd = $scanPos
                                            break
                                        }
                                    }
                                }
                            }
                        }
                        if ($fromEnd -ne $text.Length -and $fromEnd -eq $scanPos) { break }
                    }
                    $scanPos++
                }

                $fromClause = $text.Substring($fromBodyStart, $fromEnd - $fromBodyStart).Trim()

                if ($selectClause -eq '' -or $fromClause -eq '') { continue }

                $tables = Get-FromTables -FromClause $fromClause -ExcludeNames $cteNames
                $columns = Get-SelectColumns -SelectClause $selectClause

                foreach ($table in $tables) {
                    if ($columns.Count -eq 0 -or ($columns.Count -eq 1 -and $columns[0] -eq "*")) {
                        [void]$results.Add(@{
                            TableName  = $table
                            ColumnName = "*"
                            Operation  = "R"
                        })
                    }
                    else {
                        foreach ($col in $columns) {
                            [void]$results.Add(@{
                                TableName  = $table
                                ColumnName = $col
                                Operation  = "R"
                            })
                        }
                    }
                }
            }
        }
        "UPDATE" {
            $pattern = '(?i)UPDATE\s+(?:([\w$]+)\.)?([\w$]+)\s+SET\s+([\s\S]+?)(?:\s+WHERE\b|\s*;|\s*$)'
            $m = [regex]::Matches($SqlFragment, $pattern)
            foreach ($match in $m) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                $setClause = $match.Groups[3].Value
                $columns = Get-SetColumns -SetClause $setClause

                foreach ($col in $columns) {
                    [void]$results.Add(@{
                        TableName  = $tableName
                        ColumnName = $col
                        Operation  = "U"
                    })
                }
            }
        }
        "DELETE" {
            $deleteRows = Get-DeleteCrudRows -SqlFragment $SqlFragment -CteNames $cteNames
            foreach ($dr in $deleteRows) {
                [void]$results.Add($dr)
            }
        }
        "MERGE" {
            $pattern = '(?i)MERGE\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)'
            $m = [regex]::Matches($SqlFragment, $pattern)
            foreach ($match in $m) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                [void]$results.Add(@{
                    TableName  = $tableName
                    ColumnName = "(ALL)"
                    Operation  = "C"
                })
                [void]$results.Add(@{
                    TableName  = $tableName
                    ColumnName = "(ALL)"
                    Operation  = "U"
                })
            }
        }
    }

    return $results
}

function Split-ByCommaRespectingParens {
    param([string]$Text)

    $parts = [System.Collections.ArrayList]::new()
    $depth = 0
    $current = ""

    foreach ($char in $Text.ToCharArray()) {
        if ($char -eq '(') { $depth++ }
        elseif ($char -eq ')') { $depth-- }

        if ($char -eq ',' -and $depth -eq 0) {
            [void]$parts.Add($current)
            $current = ""
        }
        else {
            $current += $char
        }
    }
    if ($current.Trim() -ne '') { [void]$parts.Add($current) }

    return $parts
}

function Test-SqlFunction {
    param([string]$Name)

    $sqlFunctions = @(
        'ABS','ACOS','ADD_MONTHS','ASCII','ASIN','ATAN','ATAN2','AVG',
        'BFILENAME','BITAND',
        'CARDINALITY','CAST','CEIL','CHR','COALESCE','COLLECT','CONCAT','CONVERT',
        'CORR','COS','COSH','COUNT','CUBE','CUME_DIST','CURRENT_DATE','CURRENT_TIMESTAMP',
        'DBTIMEZONE','DECODE','DENSE_RANK','DEREF','DUMP',
        'EMPTY_BLOB','EMPTY_CLOB','EXISTSNODE','EXP','EXTRACT',
        'FIRST','FIRST_VALUE','FLOOR',
        'GREATEST','GROUPING','GROUPING_ID',
        'HEXTORAW',
        'INITCAP','INSTR','INSTRB',
        'LAG','LAST','LAST_DAY','LAST_VALUE','LEAD','LEAST','LENGTH','LENGTHB',
        'LISTAGG','LN','LNNVL','LOCALTIMESTAMP','LOG','LOWER','LPAD','LTRIM',
        'MAX','MEDIAN','MIN','MOD','MONTHS_BETWEEN',
        'NANVL','NEW_TIME','NEXT_DAY','NLS_CHARSET_ID','NLS_INITCAP','NLS_LOWER','NLS_UPPER',
        'NLSSORT','NTILE','NULLIF','NULLS','NUMTODSINTERVAL','NUMTOYMINTERVAL','NVL','NVL2',
        'OVER',
        'PERCENTILE_CONT','PERCENTILE_DISC','PERCENT_RANK','POWER',
        'RANK','RAWTOHEX','RAWTONHEX','REF','REFTOHEX','REGEXP_COUNT','REGEXP_INSTR',
        'REGEXP_LIKE','REGEXP_REPLACE','REGEXP_SUBSTR','REMAINDER','REPLACE','ROLLUP',
        'ROUND','ROW_NUMBER','ROWNUM','RPAD','RTRIM',
        'SIGN','SIN','SINH','SOUNDEX','SQRT','STDDEV','STDDEV_POP','STDDEV_SAMP',
        'SUBSTR','SUBSTRB','SUM','SYSDATE','SYSTIMESTAMP',
        'TAN','TANH','TO_CHAR','TO_CLOB','TO_DATE','TO_DSINTERVAL','TO_LOB',
        'TO_MULTI_BYTE','TO_NCHAR','TO_NCLOB','TO_NUMBER','TO_SINGLE_BYTE',
        'TO_TIMESTAMP','TO_TIMESTAMP_TZ','TO_YMINTERVAL',
        'TRANSLATE','TREAT','TRIM','TRUNC',
        'UID','UPPER','USER','USERENV',
        'VALUE','VAR_POP','VAR_SAMP','VARIANCE','VSIZE',
        'WIDTH_BUCKET',
        'XMLAGG','XMLELEMENT','XMLFOREST','XMLROOT','XMLSERIALIZE'
    )

    return ($Name.ToUpper() -in $sqlFunctions)
}

function Get-FromTables {
    param(
        [string]$FromClause,
        [System.Collections.Generic.HashSet[string]]$ExcludeNames = $null
    )

    $tables = [System.Collections.ArrayList]::new()

    $cleaned = $FromClause -replace '(?i)\b(INNER|LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL)\s+JOIN\b', ','
    $cleaned = $cleaned -replace '(?i)\bJOIN\b', ','

    $parts = Split-ByCommaRespectingParens -Text $cleaned

    $sqlKeywords = @('SELECT', 'WHERE', 'ORDER', 'GROUP', 'HAVING', 'SET', 'VALUES',
                     'INTO', 'FROM', 'AND', 'OR', 'NOT', 'NULL', 'AS', 'ON', 'IN',
                     'EXISTS', 'BETWEEN', 'LIKE', 'IS', 'CASE', 'WHEN', 'THEN',
                     'ELSE', 'END', 'BY', 'ASC', 'DESC', 'DISTINCT', 'ALL', 'ANY',
                     'LATERAL', 'CONNECT', 'START', 'WITH', 'PRIOR',
                     'UNION', 'INTERSECT', 'MINUS', 'FETCH', 'OFFSET', 'ONLY',
                     'PARTITION', 'OVER', 'ROWS', 'RANGE', 'UNBOUNDED', 'PRECEDING',
                     'FOLLOWING', 'CURRENT')

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }
        if ($trimmed.StartsWith('(')) { continue }

        $trimmed = $trimmed -replace '(?i)\bON\b.*$', ''
        $trimmed = $trimmed.Trim()
        if ($trimmed -eq '') { continue }

        if ($trimmed -match '(?:([\w$]+)\.)?([\w$]+)(?:\s+([\w$]+))?') {
            $tblName = $Matches[2].ToUpper()
            $isKeyword = $tblName -in $sqlKeywords
            $isNumber = $tblName -match '^\d+$'
            $isFunc = Test-SqlFunction -Name $tblName
            $isCte = $null -ne $ExcludeNames -and $ExcludeNames.Contains($tblName)
            if (-not $isKeyword -and -not $isNumber -and -not $isFunc -and -not $isCte) {
                [void]$tables.Add($tblName)
            }
        }
    }

    return $tables
}

function Get-SelectColumns {
    param([string]$SelectClause)

    $columns = [System.Collections.ArrayList]::new()

    if ($SelectClause.Trim() -eq '*') {
        [void]$columns.Add("*")
        return $columns
    }

    $cleaned = $SelectClause -replace '(?i)\bDISTINCT\b', ''
    $cleaned = $cleaned -replace '(?i)\bINTO\b.*$', ''
    $cleaned = $cleaned.Trim()
    $parts = Split-ByCommaRespectingParens -Text $cleaned

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }

        $alias = $null
        if ($trimmed -match '(?i)\bAS\s+(\w+)\s*$') {
            $alias = $Matches[1].ToUpper()
            $colExpr = ($trimmed -replace '(?i)\s+AS\s+\w+\s*$', '').Trim()
        }
        elseif ($trimmed -match '\s+(\w+)\s*$') {
            $candidate = $Matches[1].ToUpper()
            $exprPart = ($trimmed -replace '\s+\w+\s*$', '').Trim()
            $isReserved = $candidate -in @('SELECT','FROM','WHERE','AND','OR','NOT','NULL','CASE','WHEN','THEN','ELSE','END','BY','ASC','DESC','INTO','AS')
            if ($exprPart -ne '' -and $candidate -notmatch '^\d' -and -not $isReserved -and -not (Test-SqlFunction -Name $candidate)) {
                $alias = $candidate
                $colExpr = $exprPart
            }
            else {
                $colExpr = $trimmed
            }
        }
        else {
            $colExpr = $trimmed
        }

        if ($colExpr -match '(?i)\b(?:COUNT|SUM|AVG|MIN|MAX)\s*\(\s*\*\s*\)') {
            [void]$columns.Add("*")
            continue
        }

        if ($colExpr -match '(?i)\b(?:COUNT|SUM|AVG|MIN|MAX)\s*\(\s*(?:DISTINCT\s+)?([^)]+)\s*\)') {
            $innerAgg = $Matches[1].Trim() -replace '(?i)^DISTINCT\s+', ''
            if ($innerAgg -eq '*') {
                [void]$columns.Add("*")
                continue
            }
            if ($innerAgg -match '^\d+$') {
                continue
            }
            if ($innerAgg -match '(?:([\w$]+)\.)?([\w$]+)\s*$') {
                $innerName = $Matches[2]
                $isReservedCol = $innerName.ToUpper() -in @('NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END')
                if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol) {
                    [void]$columns.Add($innerName.ToUpper())
                }
            }
            continue
        }

        $hasParens = $colExpr.Contains('(')
        $colExprBase = $colExpr -replace '(?:\w+\.)', ''
        $isFuncExpr = $hasParens -or (Test-SqlFunction -Name $colExprBase)
        if ($isFuncExpr) {
            $funcContent = $colExpr
            while ($funcContent -match '^\w+\s*[(]') {
                $funcContent = $funcContent -replace '^\w+\s*[(]\s*', ''
                $funcContent = $funcContent -replace '\s*[)]\s*$', ''
            }
            $funcContent = $funcContent -replace '(?i)^\s*DISTINCT\s+', ''
            $funcContent = $funcContent.Trim()
            if ($funcContent -eq '*') {
                [void]$columns.Add("*")
            }
            else {
                $colPattern = '(?:(\w+)\.)?(\w+)'
                if ($funcContent -match $colPattern) {
                    $innerName = $Matches[2]
                    $isReservedCol = $innerName.ToUpper() -in @('NULL','CASE','WHEN','THEN','ELSE','END')
                    if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol -and $innerName -notmatch '^\d') {
                        [void]$columns.Add($innerName.ToUpper())
                    }
                }
            }
        }
        elseif ($colExpr -match '(?:\w+\.)?(\w+)$') {
            $colName = $Matches[1].ToUpper()
            $isReservedCol2 = $colName -in @('SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'INTO')
            if (-not $isReservedCol2 -and $colName -notmatch '^\d+$' -and -not (Test-SqlFunction -Name $colName)) {
                [void]$columns.Add($colName)
            }
        }
    }

    if ($columns.Count -eq 0 -and $SelectClause -match '(?i)\(\s*\*\s*\)') {
        [void]$columns.Add("*")
    }

    return $columns
}

function Get-SetColumns {
    param([string]$SetClause)

    $columns = [System.Collections.ArrayList]::new()
    $parts = Split-ByCommaRespectingParens -Text $SetClause

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -match '^(\w+)\s*=') {
            [void]$columns.Add($Matches[1].ToUpper())
        }
    }

    return $columns
}

function Get-PackageBodySection {
    param([string]$Content)

    $bodyPattern = '(?i)CREATE\s+OR\s+REPLACE\s+PACKAGE\s+BODY\s+'
    $bodyMatch = [regex]::Match($Content, $bodyPattern)

    if ($bodyMatch.Success) {
        return $Content.Substring($bodyMatch.Index)
    }

    return $Content
}

function ConvertFrom-OracleSqlFile {
    param(
        [string]$FilePath,
        [switch]$DebugLog
    )

    $rawContent = Get-Content $FilePath -Raw -Encoding Default
    $content = Remove-SqlComments -Content $rawContent
    $fileName = [System.IO.Path]::GetFileName($FilePath)

    $objectInfo = Get-OracleObjectInfo -Content $content -FileName $fileName

    $parseContent = $content
    if ($objectInfo.ObjectType -eq "PACKAGE") {
        $parseContent = Get-PackageBodySection -Content $content
    }

    $blocks = Get-ProcedureBlocks -Content $parseContent

    $results = [System.Collections.ArrayList]::new()

    foreach ($block in $blocks) {
        $featureName = "$($objectInfo.ObjectType):$($objectInfo.ObjectName).$($block.Name)"
        $blockExtractCount = 0

        foreach ($opType in @("INSERT", "SELECT", "UPDATE", "DELETE", "MERGE")) {
            $extracted = Get-TableAndColumns -SqlFragment $block.Content -OperationType $opType
            $blockExtractCount += $extracted.Count

            foreach ($item in $extracted) {
                [void]$results.Add(@{
                    SourceType  = "Oracle"
                    SourceFile  = $fileName
                    ObjectType  = $objectInfo.ObjectType
                    ObjectName  = $objectInfo.ObjectName
                    ProcName    = $block.Name
                    FeatureName = $featureName
                    TableName   = $item.TableName
                    ColumnName  = $item.ColumnName
                    Operation   = $item.Operation
                })
            }
        }

        if ($DebugLog) {
            $hint = if ($blockExtractCount -eq 0) { " (パーサで0件→SQL未検出の可能性)" } else { "" }
            Write-Host "[Oracle][Debug] 解析: $fileName | $($block.Name) | 抽出=$blockExtractCount | 本文=$($block.Content.Length) 文字$hint" -ForegroundColor DarkCyan
        }
    }

    return $results
}

function ConvertFrom-OracleSqlDirectory {
    param(
        [string]$SourcePath,
        [string]$FilePattern = "*.sql",
        [string[]]$ExcludePatterns = @(),
        [string[]]$ExcludeTables = @(),
        [string[]]$ExcludeSchemas = @(),
        [switch]$DebugLog
    )

    Write-Host "[Oracle] 解析開始: $SourcePath" -ForegroundColor Cyan
    if ($DebugLog) {
        Write-Host "[Oracle][Debug] 除外テーブル: $($ExcludeTables -join ', ')" -ForegroundColor DarkCyan
        Write-Host "[Oracle][Debug] 除外スキーマ: $($ExcludeSchemas -join ', ')" -ForegroundColor DarkCyan
    }

    $files = Get-ChildItem -Path $SourcePath -Filter $FilePattern -Recurse -File
    foreach ($pattern in $ExcludePatterns) {
        $files = $files | Where-Object { $_.Name -notlike $pattern }
    }

    Write-Host "[Oracle] 対象ファイル数: $($files.Count)" -ForegroundColor Cyan

    $allResults = [System.Collections.ArrayList]::new()
    $fileCount = 0

    foreach ($file in $files) {
        $fileCount++
        Write-Progress -Activity "Oracle SQL 解析中" -Status "$fileCount / $($files.Count): $($file.Name)" -PercentComplete (($fileCount / $files.Count) * 100)

        try {
            $fileResults = ConvertFrom-OracleSqlFile -FilePath $file.FullName -DebugLog:$DebugLog

            if ($DebugLog -and $fileResults.Count -gt 0) {
                $byFeature = $fileResults | Group-Object FeatureName
                foreach ($gf in $byFeature) {
                    $kept = 0
                    $dropped = 0
                    $dropReasons = [System.Collections.ArrayList]::new()
                    foreach ($r in $gf.Group) {
                        $skip = $false
                        $reason = ""
                        if ($r.TableName -in $ExcludeTables) {
                            $skip = $true
                            $reason = "ExcludeTables"
                        }
                        if (-not $skip) {
                            foreach ($schema in $ExcludeSchemas) {
                                if ($r.TableName -like "$schema*") {
                                    $skip = $true
                                    $reason = "ExcludeSchemas:$schema"
                                    break
                                }
                            }
                        }
                        if ($skip) {
                            $dropped++
                            if ($dropReasons.Count -lt 5 -and $reason -ne "") {
                                [void]$dropReasons.Add("$($r.TableName)($reason)")
                            }
                        }
                        else {
                            $kept++
                        }
                    }
                    if ($dropped -gt 0 -and $kept -eq 0) {
                        $sample = if ($dropReasons.Count -gt 0) { " 例: $($dropReasons -join ', ')" } else { "" }
                        Write-Host "[Oracle][Debug] 除外のみ(全件フィルタ): $($gf.Name) | 件数=$dropped$sample" -ForegroundColor Yellow
                    }
                    elseif ($dropped -gt 0 -and $kept -gt 0) {
                        Write-Host "[Oracle][Debug] 一部除外: $($gf.Name) | 残=$kept / 除外=$dropped" -ForegroundColor DarkCyan
                    }
                }
            }

            foreach ($result in $fileResults) {
                $skip = $false
                if ($result.TableName -in $ExcludeTables) { $skip = $true }
                foreach ($schema in $ExcludeSchemas) {
                    if ($result.TableName -like "$schema*") { $skip = $true }
                }
                if (-not $skip) {
                    [void]$allResults.Add($result)
                }
            }
        }
        catch {
            Write-Warning "[Oracle] 解析エラー: $($file.FullName) - $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "Oracle SQL 解析中" -Completed
    Write-Host "[Oracle] 解析完了: $($allResults.Count) 件のCRUDエントリを検出" -ForegroundColor Green

    return $allResults
}
