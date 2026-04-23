<#
.SYNOPSIS
    Oracle SQL ソースファイルを解析し、CRUD操作情報を抽出する

.DESCRIPTION
    Oracle PL/SQL パッケージ・トリガー・ビュー等のSQLファイルを読み込み、
    オブジェクト単位（パッケージ本文全体・1ファイル全体など）で
    INSERT/SELECT/UPDATE/DELETE/MERGE文からテーブル名・項目名・操作種別を抽出する

.PARAMETER SourcePath
    SQLファイル格納ディレクトリ

.PARAMETER FilePattern
    対象ファイルパターン（デフォルト: *.sql）

.PARAMETER ExcludePatterns
    除外ファイルパターン配列
#>

$script:CrudParseProfileStats = $null

function Reset-CrudParseProfileStats {
    if ($env:CRUD_MATRIX_PARSE_PROFILE -ne '1') { return }
    $script:CrudParseProfileStats = @{
        GetTableAndColumnsCalls           = 0
        SelectKeywordIterations           = 0
        GetColumnRefsFromPredicateText    = 0
        GetSelectColumnRefsScalarExtra    = 0
        GetOracleSqlTailLengthCalls       = 0
    }
}

function Bump-CrudParseProfile {
    param([Parameter(Mandatory)][string]$Name)
    if ($env:CRUD_MATRIX_PARSE_PROFILE -ne '1') { return }
    if ($null -eq $script:CrudParseProfileStats) { return }
    $script:CrudParseProfileStats[$Name] = [int]$script:CrudParseProfileStats[$Name] + 1
}

function Write-CrudParseProfileReport {
    param([string]$Label)
    if ($env:CRUD_MATRIX_PARSE_PROFILE -ne '1') { return }
    if ($null -eq $script:CrudParseProfileStats) { return }
    Write-Host "[ParseProfile] $Label" -ForegroundColor Cyan
    foreach ($k in ($script:CrudParseProfileStats.Keys | Sort-Object)) {
        Write-Host "  ${k}: $($script:CrudParseProfileStats[$k])" -ForegroundColor DarkCyan
    }
}

function Remove-SqlComments {
    param([string]$Content)

    $result = $Content -replace [char]0x3000, ' '
    $len = $result.Length
    $sb = [System.Text.StringBuilder]::new()
    $inString = $false
    $i = 0
    while ($i -lt $len) {
        $ch = $result[$i]
        if ($inString) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $len -and $result[$i + 1] -eq [char]0x27) {
                [void]$sb.Append($ch)
                [void]$sb.Append($result[$i + 1])
                $i += 2
                continue
            }
            if ($ch -eq [char]0x27) {
                $inString = $false
            }
            [void]$sb.Append($ch)
            $i++
            continue
        }
        if ($ch -eq [char]0x27) {
            $inString = $true
            [void]$sb.Append($ch)
            $i++
            continue
        }
        if ($ch -eq [char]0x2D -and $i + 1 -lt $len -and $result[$i + 1] -eq [char]0x2D) {
            $i += 2
            while ($i -lt $len -and $result[$i] -ne [char]10 -and $result[$i] -ne [char]13) {
                $i++
            }
            continue
        }
        if ($ch -eq [char]0x2F -and $i + 1 -lt $len -and $result[$i + 1] -eq [char]0x2A) {
            $i += 2
            while ($i + 1 -lt $len -and -not ($result[$i] -eq [char]0x2A -and $result[$i + 1] -eq [char]0x2F)) {
                $i++
            }
            if ($i + 1 -lt $len) {
                $i += 2
            }
            else {
                $i = $len
            }
            continue
        }
        [void]$sb.Append($ch)
        $i++
    }
    return $sb.ToString()
}

function Mask-OracleSqlStringLiteralsForParse {
    param([string]$Content)

    $len = $Content.Length
    $sb = [System.Text.StringBuilder]::new()
    $inString = $false
    $i = 0
    while ($i -lt $len) {
        $ch = $Content[$i]
        if ($inString) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $len -and $Content[$i + 1] -eq [char]0x27) {
                [void]$sb.Append([char]0x27)
                [void]$sb.Append([char]0x27)
                $i += 2
                continue
            }
            if ($ch -eq [char]0x27) {
                $inString = $false
                [void]$sb.Append([char]0x27)
                $i++
                continue
            }
            [void]$sb.Append(' ')
            $i++
            continue
        }
        if ($ch -eq [char]0x27) {
            $inString = $true
            [void]$sb.Append([char]0x27)
            $i++
            continue
        }
        [void]$sb.Append($ch)
        $i++
    }
    return $sb.ToString()
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

function Get-OraclePlSqlDeclaredVariableNames {
    param([string]$PlSqlBlock)

    $names = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ([string]::IsNullOrWhiteSpace($PlSqlBlock)) {
        return $names
    }
    $recordTypeNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($m in [regex]::Matches($PlSqlBlock, '(?im)\bTYPE\s+([\w\$]+)\s+IS\s+RECORD\b')) {
        [void]$recordTypeNames.Add($m.Groups[1].Value.ToUpper())
    }
    if ($recordTypeNames.Count -gt 0) {
        $skipFirstAsVar = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($w in @('TYPE', 'SUBTYPE', 'DECLARE', 'BEGIN', 'END', 'EXCEPTION', 'WHEN', 'IF', 'THEN', 'ELSE', 'ELSIF', 'LOOP', 'WHILE', 'FOR', 'OPEN', 'CLOSE', 'FETCH', 'EXIT', 'CONTINUE', 'GOTO', 'NULL', 'RAISE', 'RETURN', 'PRAGMA', 'FUNCTION', 'PROCEDURE', 'PACKAGE', 'CREATE', 'OR', 'REPLACE', 'BODY', 'USING', 'CURSOR', 'RECORD')) {
            [void]$skipFirstAsVar.Add($w)
        }
        foreach ($m in [regex]::Matches($PlSqlBlock, '(?im)^\s*([\w\$]+)\s+([\w\$]+)\s*(?:;|:=)')) {
            $firstU = $m.Groups[1].Value.ToUpper()
            $secondU = $m.Groups[2].Value.ToUpper()
            if (-not $recordTypeNames.Contains($secondU)) { continue }
            if ($skipFirstAsVar.Contains($firstU)) { continue }
            [void]$names.Add($firstU)
        }
    }
    foreach ($m in [regex]::Matches($PlSqlBlock, '(?im)^\s*([\w\$]+)\s+[\w\$\.]+%(?:TYPE|ROWTYPE)\b')) {
        [void]$names.Add($m.Groups[1].Value.ToUpper())
    }
    foreach ($m in [regex]::Matches($PlSqlBlock, '(?im)^\s*([\w\$]+)\s+(?:NUMBER|BINARY_FLOAT|BINARY_DOUBLE|INTEGER|PLS_INTEGER|BOOLEAN|DATE|TIMESTAMP|VARCHAR2|CHAR|NVARCHAR2|NCHAR|CLOB|BLOB|NCLOB)\b')) {
        [void]$names.Add($m.Groups[1].Value.ToUpper())
    }
    foreach ($m in [regex]::Matches($PlSqlBlock, '(?is)\bFOR\s+([\w\$]+)\s+IN\s+')) {
        $tail = $PlSqlBlock.Substring($m.Index + $m.Length).TrimStart()
        if ($tail -match '^(?i)REVERSE\s+') {
            $tail = ($tail -replace '^(?i)REVERSE\s+', '').TrimStart()
        }
        if ($tail.Length -gt 0 -and [char]::IsDigit($tail, 0)) {
            continue
        }
        [void]$names.Add($m.Groups[1].Value.ToUpper())
    }
    return $names
}

function Mask-OracleExistsSubqueriesInPredicateText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }
    $t = $Text
    $sb = [char[]]::new($t.Length)
    for ($i = 0; $i -lt $t.Length; $i++) {
        $sb[$i] = $t[$i]
    }
    $searchPos = 0
    while ($searchPos -lt $t.Length) {
        $sub = $t.Substring($searchPos)
        $m = [regex]::Match($sub, '(?i)\b(?:NOT\s+)?EXISTS\s*\(')
        if (-not $m.Success) {
            break
        }
        $openParen = $searchPos + $m.Index + $m.Length - 1
        $depth = 0
        $inString = $false
        $j = $openParen
        while ($j -lt $t.Length) {
            $adv = Step-OracleSqlScanOneChar -Text $t -ScanPos $j -InString ([ref]$inString) -Depth ([ref]$depth)
            $j += $adv
            if ($depth -eq 0 -and $j -gt $openParen) {
                break
            }
        }
        $maskStart = $searchPos + $m.Index
        $maskEnd = $j
        for ($k = $maskStart; $k -lt $maskEnd; $k++) {
            $sb[$k] = ' '
        }
        $searchPos = $j
    }
    # IN (SELECT ...) / NOT IN (SELECT ...) もマスクする（括弧内のみ）
    $searchPos = 0
    while ($searchPos -lt $t.Length) {
        $sub = $t.Substring($searchPos)
        $m = [regex]::Match($sub, '(?i)\b(?:NOT\s+)?IN\s*\(')
        if (-not $m.Success) {
            break
        }
        $openParen = $searchPos + $m.Index + $m.Length - 1
        # 括弧内が SELECT で始まる場合のみマスク対象
        $innerStart = $openParen + 1
        if ($innerStart -lt $t.Length) {
            $innerHead = $t.Substring($innerStart).TrimStart()
            if ($innerHead -notmatch '(?i)^SELECT\b') {
                $searchPos = $searchPos + $m.Index + 1
                continue
            }
        }
        $depth = 0
        $inString = $false
        $j = $openParen
        while ($j -lt $t.Length) {
            $adv = Step-OracleSqlScanOneChar -Text $t -ScanPos $j -InString ([ref]$inString) -Depth ([ref]$depth)
            $j += $adv
            if ($depth -eq 0 -and $j -gt $openParen) {
                break
            }
        }
        # ( から ) までをマスク（IN キーワード自体は残す）
        for ($k = $openParen; $k -lt $j; $k++) {
            $sb[$k] = ' '
        }
        $searchPos = $j
    }
    return [string]::new($sb)
}

function Get-DeleteCrudRows {
    param(
        [string]$SqlFragment,
        [System.Collections.Generic.HashSet[string]]$CteNames,
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames = $null,
        [string[]]$AdditionalCteNames = @()
    )

    $out = [System.Collections.ArrayList]::new()

    $pattern = '(?i)\bDELETE\b\s+(?:FROM\s+)?(?:([\w$]+)\.)?([\w$]+)(?:\s+(?!WHERE\b)([\w$]+))?'
    foreach ($match in [regex]::Matches($SqlFragment, $pattern)) {
        $tableName = $match.Groups[2].Value.ToUpper()
        if ($tableName -eq '') { continue }
        if ($CteNames.Count -gt 0 -and $CteNames.Contains($tableName)) { continue }
        $aliasTok = if ($match.Groups[3].Success -and $match.Groups[3].Value -ne '') { $match.Groups[3].Value.ToUpper() } else { '' }
        [void]$out.Add(@{
            TableName  = $tableName
            ColumnName = "(ALL)"
            Operation  = "D"
        })
        $afterMatch = $SqlFragment.Substring($match.Index + $match.Length)
        if ($afterMatch -match '(?is)^\s*WHERE\b') {
            $wm = [regex]::Match($afterMatch, '(?is)^\s*WHERE\b([\s\S]+)$')
            if ($wm.Success) {
                $whereRaw = $wm.Groups[1].Value
                $semiPos = $whereRaw.IndexOf(';')
                if ($semiPos -ge 0) {
                    $whereText = $whereRaw.Substring(0, $semiPos).Trim()
                }
                else {
                    $whereText = $whereRaw.Trim().TrimEnd(';')
                }
                $whereForOuterRefs = Mask-OracleExistsSubqueriesInPredicateText -Text $whereText
                $whereRefs = Get-ColumnRefsFromPredicateText -Text $whereForOuterRefs
                $aliasMap = @{}
                $aliasMap[$tableName] = $tableName
                if ($aliasTok -ne '') {
                    $aliasMap[$aliasTok] = $tableName
                }
                foreach ($wr in $whereRefs) {
                    $qual = if ($null -ne $wr.TableName -and $wr.TableName -ne '') { $wr.TableName.ToUpper() } else { '' }
                    if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                        continue
                    }
                    $resolvedTable = $null
                    if ($qual -ne '') {
                        if ($aliasMap.ContainsKey($qual)) {
                            $resolvedTable = $tableName
                        }
                        else {
                            $resolvedTable = $qual
                        }
                    }
                    else {
                        $resolvedTable = $tableName
                    }
                    [void]$out.Add(@{
                        TableName  = $resolvedTable
                        ColumnName = $wr.ColumnName
                        Operation  = "R"
                    })
                }
                foreach ($er in (Get-CrudRowsFromExistsSubqueriesInText -Text $whereText -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)) {
                    $dup = $false
                    foreach ($existing in $out) {
                        if ($null -ne $existing.TableName -and $existing.TableName -eq $er.TableName -and $existing.ColumnName -eq $er.ColumnName -and $existing.Operation -eq $er.Operation) {
                            $dup = $true
                            break
                        }
                    }
                    if (-not $dup) {
                        [void]$out.Add($er)
                    }
                }
            }
        }
    }

    return $out
}

function Step-OracleSqlScanOneChar {
    param(
        [string]$Text,
        [int]$ScanPos,
        [ref]$InString,
        [ref]$Depth
    )

    $ch = $Text[$ScanPos]
    if ($InString.Value) {
        if ($ch -eq [char]0x27 -and $ScanPos + 1 -lt $Text.Length -and $Text[$ScanPos + 1] -eq [char]0x27) {
            return 2
        }
        if ($ch -eq [char]0x27) {
            $InString.Value = $false
        }
        return 1
    }
    if ($ch -eq [char]0x27) {
        $InString.Value = $true
        return 1
    }
    if ($ch -eq '(') {
        $Depth.Value++
    }
    elseif ($ch -eq ')') {
        $Depth.Value--
    }
    return 1
}

function Get-OracleSqlTailLengthToSemicolonAtDepthZero {
    param(
        [string]$Text,
        [int]$StartPos
    )

    Bump-CrudParseProfile -Name 'GetOracleSqlTailLengthCalls'

    if ($null -eq $Text -or $Text.Length -eq 0 -or $StartPos -lt 0 -or $StartPos -ge $Text.Length) {
        return 0
    }
    $depth = 0
    $inString = $false
    $i = $StartPos
    while ($i -lt $Text.Length) {
        if (-not $inString -and $depth -eq 0 -and $Text[$i] -eq ';') {
            return $i - $StartPos + 1
        }
        $adv = Step-OracleSqlScanOneChar -Text $Text -ScanPos $i -InString ([ref]$inString) -Depth ([ref]$depth)
        $i += $adv
    }
    return $Text.Length - $StartPos
}

function Remove-OracleWindowOverClauses {
    param([string]$Text)

    if ($null -eq $Text -or $Text.Trim() -eq '') {
        return $Text
    }
    $sb = New-Object System.Text.StringBuilder
    $len = $Text.Length
    $i = 0
    $inStr = $false
    while ($i -lt $len) {
        $ch = $Text[$i]
        if ($inStr) {
            [void]$sb.Append($ch)
            if ($ch -eq [char]0x27 -and $i + 1 -lt $len -and $Text[$i + 1] -eq [char]0x27) {
                [void]$sb.Append($Text[$i + 1])
                $i += 2
                continue
            }
            if ($ch -eq [char]0x27) { $inStr = $false }
            $i++
            continue
        }
        if ($ch -eq [char]0x27) {
            $inStr = $true
            [void]$sb.Append($ch)
            $i++
            continue
        }
        $tail = $Text.Substring($i)
        $m = [regex]::Match($tail, '(?i)^\bOVER\s*\(')
        if ($m.Success) {
            $openParen = $i + $m.Length - 1
            $depth = 0
            $inStr2 = $false
            $j = $openParen
            while ($j -lt $len) {
                $adv = Step-OracleSqlScanOneChar -Text $Text -ScanPos $j -InString ([ref]$inStr2) -Depth ([ref]$depth)
                $j += $adv
                if ($depth -eq 0 -and $j -gt $openParen) { break }
            }
            $i = $j
            [void]$sb.Append(' ')
            continue
        }
        [void]$sb.Append($ch)
        $i++
    }
    return $sb.ToString()
}

function Test-OracleFromClauseKeywordAt {
    param(
        [string]$Text,
        [int]$ScanPos
    )

    if ($ScanPos -lt 0 -or $ScanPos + 3 -ge $Text.Length) {
        return $false
    }
    $tailLen = [Math]::Min(8, $Text.Length - $ScanPos)
    if ($tailLen -lt 4) {
        return $false
    }
    if ($Text.Substring($ScanPos, $tailLen) -notmatch '(?i)^FROM\b') {
        return $false
    }
    if ($ScanPos -gt 0) {
        $prev = $Text[$ScanPos - 1]
        if ([char]::IsLetterOrDigit($prev) -or $prev -eq '_' -or $prev -eq '$' -or $prev -eq '#') {
            return $false
        }
    }
    return $true
}

function Test-OracleWhereClauseKeywordAt {
    param(
        [string]$Text,
        [int]$ScanPos
    )

    if ($ScanPos -lt 0 -or $ScanPos + 4 -ge $Text.Length) {
        return $false
    }
    $tailLen = [Math]::Min(12, $Text.Length - $ScanPos)
    if ($tailLen -lt 5) {
        return $false
    }
    if ($Text.Substring($ScanPos, $tailLen) -notmatch '(?i)^WHERE\b') {
        return $false
    }
    if ($ScanPos -gt 0) {
        $prev = $Text[$ScanPos - 1]
        if ([char]::IsLetterOrDigit($prev) -or $prev -eq '_' -or $prev -eq '$' -or $prev -eq '#') {
            return $false
        }
    }
    return $true
}

function Split-OracleUpdateTailSetAndWhere {
    param([string]$Tail)

    if ([string]::IsNullOrWhiteSpace($Tail)) {
        return @{ SetClause = ''; WhereText = $null }
    }
    $depth = 0
    $inStr = $false
    $whereIdx = -1
    for ($i = 0; $i -lt $Tail.Length; $i++) {
        $ch = $Tail[$i]
        if ($inStr) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $Tail.Length -and $Tail[$i + 1] -eq [char]0x27) {
                $i++
                continue
            }
            if ($ch -eq [char]0x27) {
                $inStr = $false
            }
            continue
        }
        if ($ch -eq [char]0x27) {
            $inStr = $true
            continue
        }
        if ($ch -eq '(') {
            $depth++
        }
        elseif ($ch -eq ')') {
            $depth--
            if ($depth -lt 0) {
                $depth = 0
            }
        }
        elseif ($depth -eq 0 -and (Test-OracleWhereClauseKeywordAt -Text $Tail -ScanPos $i)) {
            $whereIdx = $i
            break
        }
    }
    if ($whereIdx -lt 0) {
        $one = ($Tail -split ';')[0].Trim()
        return @{ SetClause = $one; WhereText = $null }
    }
    $setClause = $Tail.Substring(0, $whereIdx).Trim()
    $afterWhere = $Tail.Substring($whereIdx)
    $wm = [regex]::Match($afterWhere, '(?is)^\s*WHERE\b\s*([\s\S]+)$')
    $whereText = $null
    if ($wm.Success) {
        $whereRaw = $wm.Groups[1].Value
        $semiPos = $whereRaw.IndexOf(';')
        if ($semiPos -ge 0) {
            $whereText = $whereRaw.Substring(0, $semiPos).Trim()
        }
        else {
            $whereText = $whereRaw.Trim().TrimEnd(';')
        }
    }
    return @{ SetClause = $setClause; WhereText = $whereText }
}

function Test-OracleSelectKeywordAt {
    param(
        [string]$Text,
        [int]$ScanPos
    )

    if ($ScanPos -lt 0 -or ($ScanPos + 6) -gt $Text.Length) {
        return $false
    }
    $tailLen = [Math]::Min(12, $Text.Length - $ScanPos)
    if ($Text.Substring($ScanPos, $tailLen) -notmatch '(?i)^SELECT\b') {
        return $false
    }
    if ($ScanPos -gt 0) {
        $prev = $Text[$ScanPos - 1]
        if ([char]::IsLetterOrDigit($prev) -or $prev -eq '_' -or $prev -eq '$' -or $prev -eq '#') {
            return $false
        }
    }
    return $true
}

function Test-OracleForUpdateOrShareClauseAt {
    param(
        [string]$Text,
        [int]$ScanPos
    )

    if ($ScanPos + 7 -gt $Text.Length) {
        return $false
    }
    $tailLen = [Math]::Min(32, $Text.Length - $ScanPos)
    return $Text.Substring($ScanPos, $tailLen) -match '(?i)^FOR\s+(UPDATE|SHARE)\b'
}

function Test-OracleInSingleQuotedStringAt {
    param([string]$Text, [int]$Pos)

    if ($Pos -le 0) {
        return $false
    }
    $inStr = $false
    $i = 0
    while ($i -lt $Pos) {
        if ($Text[$i] -eq [char]0x27 -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq [char]0x27) {
            $i += 2
            continue
        }
        if ($Text[$i] -eq [char]0x27) {
            $inStr = -not $inStr
        }
        $i++
    }
    return $inStr
}

function Get-OracleSqlFragmentToStatementEnd {
    param([string]$Text, [int]$StartPos)

    if ($StartPos -ge $Text.Length) {
        return ''
    }
    if (Test-OracleInSingleQuotedStringAt -Text $Text -Pos $StartPos) {
        $i = $StartPos
        while ($i -lt $Text.Length) {
            if ($Text[$i] -eq [char]0x27 -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq [char]0x27) {
                $i += 2
                continue
            }
            if ($Text[$i] -eq [char]0x27) {
                return $Text.Substring($StartPos, $i - $StartPos)
            }
            $i++
        }
        return $Text.Substring($StartPos)
    }
    $depth = 0
    $inString = $false
    $i = $StartPos
    while ($i -lt $Text.Length) {
        $ch = $Text[$i]
        if ($inString) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq [char]0x27) {
                $i += 2
                continue
            }
            if ($ch -eq [char]0x27) {
                $inString = $false
            }
            $i++
            continue
        }
        if ($ch -eq [char]0x27) {
            $inString = $true
            $i++
            continue
        }
        if ($ch -eq '(') {
            $depth++
        }
        elseif ($ch -eq ')') {
            $depth--
        }
        elseif ($ch -eq ';' -and $depth -eq 0) {
            return $Text.Substring($StartPos, $i - $StartPos)
        }
        $i++
    }
    return $Text.Substring($StartPos)
}

function Get-TableAndColumns {
    param(
        [string]$SqlFragment,
        [string]$OperationType,
        [string[]]$AdditionalCteNames = @(),
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames = $null
    )

    Bump-CrudParseProfile -Name 'GetTableAndColumnsCalls'

    $SqlFragment = $SqlFragment -replace [char]0x3000, ' '
    $SqlFragment = $SqlFragment -replace '(?is)\bOPEN\s+[\w$"]+\s+FOR\s+(?!SELECT\b|WITH\b)([\w$]+(?:\.[\w$]+)*)\s+USING\b', ' '

    $crudExtractList = [System.Collections.ArrayList]::new()

    $cteNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($SqlFragment -match '(?i)\bWITH\b') {
        $cteMatches = [regex]::Matches($SqlFragment, '(?i)([\w$]+)\s+AS\s*\(')
        foreach ($cm in $cteMatches) {
            [void]$cteNames.Add($cm.Groups[1].Value.ToUpper())
        }
    }
    foreach ($extra in $AdditionalCteNames) {
        if (-not [string]::IsNullOrWhiteSpace($extra)) {
            [void]$cteNames.Add($extra.Trim().ToUpper())
        }
    }

    switch ($OperationType) {
        "INSERT" {
            # テーブル名と列リストの間に別名があってもよい（例: INSERT INTO TBL A (COL1, A.COL2, ...)）
            $pattern = '(?i)INSERT\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)(?:\s+(?!(?:VALUES|SELECT)\b)[\w$]+)?\s*\(([^)]+)\)'
            $m = [regex]::Matches($SqlFragment, $pattern)
            foreach ($match in $m) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                $columnsRaw = $match.Groups[3].Value
                # 別名修飾子（例: A.COL2 → COL2）を除去してカラム名を正規化する
                $columns = ($columnsRaw -split ',') | ForEach-Object {
                    $t = $_.Trim().ToUpper()
                    if ($t -match '\.([\w$]+)$') { $Matches[1] } else { $t }
                } | Where-Object { $_ -ne '' }
                foreach ($col in $columns) {
                    [void]$crudExtractList.Add(@{
                        TableName  = $tableName
                        ColumnName = $col
                        Operation  = "C"
                    })
                }
            }

            $patternInsertSelect = '(?is)INSERT\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)\s+SELECT\s+'
            foreach ($match in [regex]::Matches($SqlFragment, $patternInsertSelect)) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                $tailStart = $match.Index + $match.Length
                $tail = Get-OracleSqlFragmentToStatementEnd -Text $SqlFragment -StartPos $tailStart
                if ($tail.Trim() -eq '') { continue }
                $trimTail = $tail.Trim().TrimEnd(';')
                $innerSelectSql = 'SELECT ' + $trimTail
                $selResults = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $innerSelectSql -OperationType "SELECT" -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)
                $colNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                foreach ($sr in $selResults) {
                    if ($sr.Operation -eq 'R' -and $sr.ColumnName -ne '*' -and $sr.ColumnName -ne '(ALL)') {
                        [void]$colNames.Add($sr.ColumnName)
                    }
                }
                if ($colNames.Count -eq 0) {
                    [void]$crudExtractList.Add(@{
                        TableName  = $tableName
                        ColumnName = "*"
                        Operation  = "C"
                    })
                }
                else {
                    foreach ($cn in $colNames) {
                        [void]$crudExtractList.Add(@{
                            TableName  = $tableName
                            ColumnName = $cn
                            Operation  = "C"
                        })
                    }
                }
            }

            # テーブル名と VALUES の間に別名があってもよく、VALUES の後に括弧なし変数も許容する（例: INSERT INTO TBL A VALUES d_row）
            $patternInsertValues = '(?i)INSERT\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)(?:\s+(?!VALUES\b)[\w$]+)?\s+VALUES\b'
            foreach ($im in [regex]::Matches($SqlFragment, $patternInsertValues)) {
                $tableName = $im.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                [void]$crudExtractList.Add(@{
                    TableName  = $tableName
                    ColumnName = "*"
                    Operation  = "C"
                })
            }
        }
        "SELECT" {
            $selectMatches = [regex]::Matches($SqlFragment, '(?i)\bSELECT\b')
            foreach ($sm in $selectMatches) {
                Bump-CrudParseProfile -Name 'SelectKeywordIterations'
                $selIdx = $sm.Index
                $j = $selIdx - 1
                while ($j -ge 0 -and [char]::IsWhiteSpace($SqlFragment[$j])) {
                    $j--
                }
                if ($j -ge 0 -and $SqlFragment[$j] -eq ',') {
                    continue
                }

                $depth = 0
                $inString = $false
                $selectStart = $sm.Index + $sm.Length
                $fromStart = -1
                $fromEnd = -1
                $scanPos = $selectStart
                $text = $SqlFragment
                $nestedSelectFromSkipsRemaining = 0

                while ($scanPos -lt $text.Length) {
                    if (-not $inString -and $depth -eq 0) {
                        if (Test-OracleSelectKeywordAt -Text $text -ScanPos $scanPos) {
                            $nestedSelectFromSkipsRemaining++
                        }
                        else {
                            $c0 = $text[$scanPos]
                            if ($c0 -eq 'F' -or $c0 -eq 'f') {
                                if (Test-OracleFromClauseKeywordAt -Text $text -ScanPos $scanPos) {
                                    if ($nestedSelectFromSkipsRemaining -gt 0) {
                                        $nestedSelectFromSkipsRemaining--
                                    }
                                    else {
                                        $fromStart = $scanPos
                                        break
                                    }
                                }
                            }
                        }
                    }
                    $adv = Step-OracleSqlScanOneChar -Text $text -ScanPos $scanPos -InString ([ref]$inString) -Depth ([ref]$depth)
                    if ($depth -lt 0) { break }
                    $scanPos += $adv
                }

                if ($fromStart -lt 0) { continue }

                $selectClause = $text.Substring($selectStart, $fromStart - $selectStart).Trim()
                $afterFrom = $fromStart + 4
                if ($afterFrom -ge $text.Length) { continue }

                $depth = 0
                $inString = $false
                $fromBodyStart = $afterFrom
                $fromEnd = $text.Length
                $scanPos = $afterFrom

                $terminators = @('WHERE', 'ORDER', 'GROUP', 'HAVING', 'UNION', 'INTERSECT', 'MINUS', 'FETCH', 'CONNECT', 'PIVOT', 'UNPIVOT', 'MODEL')

                while ($scanPos -lt $text.Length) {
                    $ch = $text[$scanPos]
                    if (-not $inString -and $depth -eq 0) {
                        if ($scanPos -gt 0 -and $text[$scanPos - 1] -match '\s') {
                            if (Test-OracleForUpdateOrShareClauseAt -Text $text -ScanPos $scanPos) {
                                $fromEnd = $scanPos
                            }
                        }
                        if ($fromEnd -ne $text.Length -and $fromEnd -eq $scanPos) { break }
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
                        if ($ch -eq ';') {
                            $fromEnd = $scanPos
                            break
                        }
                    }
                    $oldPos = $scanPos
                    $adv = Step-OracleSqlScanOneChar -Text $text -ScanPos $scanPos -InString ([ref]$inString) -Depth ([ref]$depth)
                    if ($depth -lt 0) {
                        $fromEnd = $oldPos
                        break
                    }
                    $scanPos += $adv
                }

                $fromClause = $text.Substring($fromBodyStart, $fromEnd - $fromBodyStart).Trim()
                $selectClause = Remove-OracleWindowOverClauses -Text $selectClause

                if ($selectClause -eq '' -or $fromClause -eq '') { continue }

                $tablesOuter = Normalize-OracleTableList (Get-FromTables -FromClause $fromClause -ExcludeNames $cteNames)
                $tablesNested = Get-FromTablesFromNestedSelectsInSelectClause -SelectClause $selectClause -ExcludeNames $cteNames -PlSqlDeclaredNames $PlSqlDeclaredNames
                $tablesMerged = [System.Collections.ArrayList]::new()
                foreach ($t in $tablesOuter) { [void]$tablesMerged.Add($t) }
                foreach ($nt in $tablesNested) {
                    $exists = $false
                    foreach ($t in $tablesMerged) {
                        if ($t -eq $nt) { $exists = $true; break }
                    }
                    if (-not $exists) { [void]$tablesMerged.Add($nt) }
                }
                $aliasToTable = Get-OracleFromAliasToTableMap -FromClause $fromClause
                $tablesResolved = Resolve-OraclePhysicalTableNamesOrdered -Tables $tablesMerged -AliasMap $aliasToTable
                $tables = Normalize-OracleTableList $tablesResolved
                $refInfo = Get-SelectColumnRefs -SelectClause $selectClause
                $afterTailLen = Get-OracleSqlTailLengthToSemicolonAtDepthZero -Text $text -StartPos $fromEnd
                $afterFromSegment = $text.Substring($fromEnd, $afterTailLen)
                $predRefs = Get-ColumnRefsFromPredicateText -Text ($fromClause + " " + $afterFromSegment)
                $tablesOuterResolved = Resolve-OraclePhysicalTableNamesOrdered -Tables $tablesOuter -AliasMap $aliasToTable
                $firstTableResolved = if ($tablesOuterResolved.Count -gt 0) { $tablesOuterResolved[0] } else { $tables[0] }

                foreach ($table in $tables) {
                    $selectStarOnly = $refInfo.StarOnly
                    $firstTable = $firstTableResolved
                    $colsForTable = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                    if (-not $selectStarOnly) {
                        foreach ($colRef in $refInfo.Refs) {
                            $qual = if ($null -ne $colRef.TableName -and $colRef.TableName -ne '') { $colRef.TableName.ToUpper() } else { '' }
                            if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                                continue
                            }
                            $resolvedPhysical = $null
                            if ($qual -ne '') {
                                if ($aliasToTable.ContainsKey($qual)) {
                                    $resolvedPhysical = $aliasToTable[$qual]
                                }
                                else {
                                    $resolvedPhysical = $qual
                                }
                            }
                            if ($null -ne $resolvedPhysical -and $resolvedPhysical -eq $table) {
                                [void]$colsForTable.Add($colRef.ColumnName)
                            }
                            elseif (($null -eq $colRef.TableName -or $colRef.TableName -eq '') -and ($table -eq $firstTable)) {
                                [void]$colsForTable.Add($colRef.ColumnName)
                            }
                        }
                    }
                    foreach ($pr in $predRefs) {
                        $qual = if ($null -ne $pr.TableName -and $pr.TableName -ne '') { $pr.TableName.ToUpper() } else { '' }
                        if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                            continue
                        }
                        $resolvedPhysical = $null
                        if ($qual -ne '') {
                            if ($aliasToTable.ContainsKey($qual)) {
                                $resolvedPhysical = $aliasToTable[$qual]
                            }
                            else {
                                $resolvedPhysical = $qual
                            }
                        }
                        if ($null -ne $resolvedPhysical -and $resolvedPhysical -eq $table) {
                            [void]$colsForTable.Add($pr.ColumnName)
                        }
                        elseif (($null -eq $pr.TableName -or $pr.TableName -eq '') -and ($table -eq $firstTable)) {
                            [void]$colsForTable.Add($pr.ColumnName)
                        }
                    }
                    if ($selectStarOnly) {
                        [void]$crudExtractList.Add(@{
                            TableName  = $table
                            ColumnName = "*"
                            Operation  = "R"
                        })
                    }
                    foreach ($col in $colsForTable) {
                        [void]$crudExtractList.Add(@{
                            TableName  = $table
                            ColumnName = $col
                            Operation  = "R"
                        })
                    }
                }
                $existsPredText = $fromClause + " " + $afterFromSegment
                foreach ($er in (Get-CrudRowsFromExistsSubqueriesInText -Text $existsPredText -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)) {
                    [void]$crudExtractList.Add($er)
                }
            }
        }
        "UPDATE" {
            # 表名の直後に別名があってもよい（例: UPDATE UPD_TBL U SET ...）
            $pattern = '(?i)\bUPDATE\b\s+(?:([\w$]+)\.)?([\w$]+)(?:\s+([\w$]+))?\s+SET\s+'
            $m = [regex]::Matches($SqlFragment, $pattern)
            foreach ($match in $m) {
                $tableName = $match.Groups[2].Value.ToUpper()
                if ($cteNames.Count -gt 0 -and $cteNames.Contains($tableName)) { continue }
                $aliasTok = if ($match.Groups[3].Success -and $match.Groups[3].Value -ne '') { $match.Groups[3].Value.ToUpper() } else { '' }
                $tail = $SqlFragment.Substring($match.Index + $match.Length)
                $swUpd = Split-OracleUpdateTailSetAndWhere -Tail $tail
                $setClause = [string]$swUpd.SetClause
                $whereText = $swUpd.WhereText
                $columns = Get-SetColumns -SetClause $setClause

                foreach ($col in $columns) {
                    [void]$crudExtractList.Add(@{
                        TableName  = $tableName
                        ColumnName = $col
                        Operation  = "U"
                    })
                }
                $updReadMap = @{}
                $updReadMap[$tableName] = $tableName
                if ($aliasTok -ne '') {
                    $updReadMap[$aliasTok] = $tableName
                }
                Add-UpdateSetRhsReadRows -SetClause $setClause -OuterAliasToTable $updReadMap -DefaultTableName $tableName `
                    -CteNames $cteNames -PlSqlDeclaredNames $PlSqlDeclaredNames -AdditionalCteNames $AdditionalCteNames -OutList $crudExtractList
                if ($null -ne $whereText -and $whereText -ne '') {
                    $whereForOuterRefs = Mask-OracleExistsSubqueriesInPredicateText -Text $whereText
                    $whereRefs = Get-ColumnRefsFromPredicateText -Text $whereForOuterRefs
                    $aliasMap = @{}
                    $aliasMap[$tableName] = $tableName
                    if ($aliasTok -ne '') {
                        $aliasMap[$aliasTok] = $tableName
                    }
                    foreach ($wr in $whereRefs) {
                        $qual = if ($null -ne $wr.TableName -and $wr.TableName -ne '') { $wr.TableName.ToUpper() } else { '' }
                        if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                            continue
                        }
                        $resolvedTable = $null
                        if ($qual -ne '') {
                            if ($aliasMap.ContainsKey($qual)) {
                                $resolvedTable = $tableName
                            }
                            else {
                                $resolvedTable = $qual
                            }
                        }
                        else {
                            $resolvedTable = $tableName
                        }
                        [void]$crudExtractList.Add(@{
                            TableName  = $resolvedTable
                            ColumnName = $wr.ColumnName
                            Operation  = "R"
                        })
                    }
                    foreach ($er in (Get-CrudRowsFromExistsSubqueriesInText -Text $whereText -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)) {
                        [void]$crudExtractList.Add(@{
                            TableName  = $er.TableName
                            ColumnName = $er.ColumnName
                            Operation  = $er.Operation
                        })
                    }
                }
            }
        }
        "DELETE" {
            $deleteRows = Get-DeleteCrudRows -SqlFragment $SqlFragment -CteNames $cteNames -PlSqlDeclaredNames $PlSqlDeclaredNames -AdditionalCteNames $AdditionalCteNames
            foreach ($dr in $deleteRows) {
                [void]$crudExtractList.Add($dr)
            }
        }
        "MERGE" {
            $mergeRows = Get-MergeCrudRowsDetailed -SqlFragment $SqlFragment -CteNames $cteNames -PlSqlDeclaredNames $PlSqlDeclaredNames -AdditionalCteNames $AdditionalCteNames
            foreach ($mr in $mergeRows) {
                $tn = $mr.TableName
                if ($null -eq $tn -or [string]::IsNullOrWhiteSpace($tn)) { continue }
                [void]$crudExtractList.Add(@{
                    TableName  = $tn
                    ColumnName = $mr.ColumnName
                    Operation  = $mr.Operation
                })
            }
        }
    }

    return $crudExtractList
}

function Get-OracleRoughFromClauseAfterFromKeyword {
    param([string]$SelectSql)

    if ([string]::IsNullOrWhiteSpace($SelectSql)) {
        return ''
    }
    $text = $SelectSql
    $selectMatches = [regex]::Matches($text, '(?i)\bSELECT\b')
    $sm = $null
    foreach ($match in $selectMatches) {
        $selIdx = $match.Index
        $j = $selIdx - 1
        while ($j -ge 0 -and [char]::IsWhiteSpace($text[$j])) {
            $j--
        }
        if ($j -ge 0 -and $text[$j] -eq ',') {
            continue
        }
        $sm = $match
        break
    }
    if ($null -eq $sm) {
        return ''
    }

    $selectStart = $sm.Index + $sm.Length
    $depth = 0
    $inString = $false
    $fromStart = -1
    $scanPos = $selectStart
    $nestedSelectFromSkipsRemaining = 0

    while ($scanPos -lt $text.Length) {
        if (-not $inString -and $depth -eq 0) {
            if (Test-OracleSelectKeywordAt -Text $text -ScanPos $scanPos) {
                $nestedSelectFromSkipsRemaining++
            }
            else {
                $c0 = $text[$scanPos]
                if ($c0 -eq 'F' -or $c0 -eq 'f') {
                    if (Test-OracleFromClauseKeywordAt -Text $text -ScanPos $scanPos) {
                        if ($nestedSelectFromSkipsRemaining -gt 0) {
                            $nestedSelectFromSkipsRemaining--
                        }
                        else {
                            $fromStart = $scanPos
                            break
                        }
                    }
                }
            }
        }
        $adv = Step-OracleSqlScanOneChar -Text $text -ScanPos $scanPos -InString ([ref]$inString) -Depth ([ref]$depth)
        if ($depth -lt 0) { break }
        $scanPos += $adv
    }

    if ($fromStart -lt 0) {
        return ''
    }

    $afterFrom = $fromStart + 4
    if ($afterFrom -ge $text.Length) {
        return ''
    }

    $depth = 0
    $inString = $false
    $fromBodyStart = $afterFrom
    $fromEnd = $text.Length
    $scanPos = $afterFrom

    $terminators = @('WHERE', 'ORDER', 'GROUP', 'HAVING', 'UNION', 'INTERSECT', 'MINUS', 'FETCH', 'CONNECT', 'PIVOT', 'UNPIVOT', 'MODEL')

    while ($scanPos -lt $text.Length) {
        $ch = $text[$scanPos]
        if (-not $inString -and $depth -eq 0) {
            if ($scanPos -gt 0 -and $text[$scanPos - 1] -match '\s') {
                if (Test-OracleForUpdateOrShareClauseAt -Text $text -ScanPos $scanPos) {
                    $fromEnd = $scanPos
                }
            }
            if ($fromEnd -ne $text.Length -and $fromEnd -eq $scanPos) { break }
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
            if ($ch -eq ';') {
                $fromEnd = $scanPos
                break
            }
        }
        $oldPos = $scanPos
        $adv = Step-OracleSqlScanOneChar -Text $text -ScanPos $scanPos -InString ([ref]$inString) -Depth ([ref]$depth)
        if ($depth -lt 0) {
            $fromEnd = $oldPos
            break
        }
        $scanPos += $adv
    }

    $rough = $text.Substring($fromBodyStart, $fromEnd - $fromBodyStart).Trim()
    return ($rough -replace '(?i)\s+AS\s+', ' ')
}

function Add-SimpleExistsSelectUnqualifiedWhereEqReads {
    param(
        [string]$InnerSelectSql,
        [System.Collections.Generic.HashSet[string]]$CteNames
    )

    $out = [System.Collections.ArrayList]::new()
    if ([string]::IsNullOrWhiteSpace($InnerSelectSql)) {
        return ,[object[]]@()
    }
    $s = $InnerSelectSql.Trim()
    if ($s -notmatch '(?is)^\s*SELECT\s+.+?\bFROM\s+(?:([\w\$]+)\.)?([\w\$]+)\s+WHERE\s+([\w\$]+)\s*=') {
        return ,[object[]]@()
    }
    $tbl = $Matches[2].ToUpper()
    $roughFrom = Get-OracleRoughFromClauseAfterFromKeyword -SelectSql $s
    if ($roughFrom -ne '') {
        $amSimple = Get-OracleFromAliasToTableMap -FromClause $roughFrom
        if ($amSimple.ContainsKey($tbl)) {
            $tbl = [string]$amSimple[$tbl]
        }
    }
    if ($null -ne $CteNames -and $CteNames.Count -gt 0 -and $CteNames.Contains($tbl)) {
        return ,[object[]]@()
    }
    $col = $Matches[3].ToUpper()
    $kwSkip = @('SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'IN', 'IS', 'LIKE', 'EXISTS', 'BETWEEN')
    if ($col -in $kwSkip) {
        return ,[object[]]@()
    }
    [void]$out.Add(@{
        TableName  = $tbl
        ColumnName = $col
        Operation  = "R"
    })
    return ,[object[]]@($out.ToArray())
}

function Get-CrudRowsFromExistsSubqueriesInText {
    param(
        [string]$Text,
        [string[]]$AdditionalCteNames = @(),
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames = $null
    )

    $out = [System.Collections.ArrayList]::new()
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ,[object[]]@()
    }

    $cteForExists = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($extra in $AdditionalCteNames) {
        if (-not [string]::IsNullOrWhiteSpace($extra)) {
            [void]$cteForExists.Add($extra.Trim().ToUpper())
        }
    }

    $t = $Text
    $searchPos = 0
    while ($searchPos -lt $t.Length) {
        $sub = $t.Substring($searchPos)
        $m = [regex]::Match($sub, '(?i)\b(?:NOT\s+)?EXISTS\s*\(')
        if (-not $m.Success) {
            break
        }
        $openParen = $searchPos + $m.Index + $m.Length - 1
        $depth = 0
        $inString = $false
        $j = $openParen
        while ($j -lt $t.Length) {
            $adv = Step-OracleSqlScanOneChar -Text $t -ScanPos $j -InString ([ref]$inString) -Depth ([ref]$depth)
            $j += $adv
            if ($depth -eq 0 -and $j -gt $openParen) {
                break
            }
        }
        if ($j -gt $openParen + 1) {
            $innerLen = $j - $openParen - 2
            if ($innerLen -gt 0) {
                $innerTrim = $t.Substring($openParen + 1, $innerLen).Trim()
                if ($innerTrim -match '(?is)^\s*SELECT\b') {
                    $subRows = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $innerTrim -OperationType "SELECT" -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)
                    $roughFromExists = Get-OracleRoughFromClauseAfterFromKeyword -SelectSql $innerTrim
                    $amExistsInner = Get-OracleFromAliasToTableMap -FromClause $roughFromExists
                    foreach ($sr in $subRows) {
                        if ($null -eq $sr -or $null -eq $sr.TableName -or $sr.TableName -eq '') { continue }
                        $tnEx = [string]$sr.TableName
                        if ($amExistsInner.ContainsKey($tnEx)) {
                            $tnEx = [string]$amExistsInner[$tnEx]
                        }
                        [void]$out.Add(@{
                            TableName  = $tnEx
                            ColumnName = $sr.ColumnName
                            Operation  = $sr.Operation
                        })
                    }
                    foreach ($sup in (Add-SimpleExistsSelectUnqualifiedWhereEqReads -InnerSelectSql $innerTrim -CteNames $cteForExists)) {
                        $dup = $false
                        foreach ($existing in $out) {
                            if ($null -ne $existing.TableName -and $existing.TableName -eq $sup.TableName -and $existing.ColumnName -eq $sup.ColumnName -and $existing.Operation -eq $sup.Operation) {
                                $dup = $true
                                break
                            }
                        }
                        if (-not $dup) {
                            [void]$out.Add($sup)
                        }
                    }
                }
            }
        }
        $searchPos = $j
    }

    # IN (SELECT ...) / NOT IN (SELECT ...) サブクエリも EXISTS と同様に処理する
    $searchPos = 0
    while ($searchPos -lt $t.Length) {
        $sub = $t.Substring($searchPos)
        $m = [regex]::Match($sub, '(?i)\b(?:NOT\s+)?IN\s*\(')
        if (-not $m.Success) {
            break
        }
        $openParen = $searchPos + $m.Index + $m.Length - 1
        $depth = 0
        $inString = $false
        $j = $openParen
        while ($j -lt $t.Length) {
            $adv = Step-OracleSqlScanOneChar -Text $t -ScanPos $j -InString ([ref]$inString) -Depth ([ref]$depth)
            $j += $adv
            if ($depth -eq 0 -and $j -gt $openParen) {
                break
            }
        }
        if ($j -gt $openParen + 1) {
            $innerLen = $j - $openParen - 2
            if ($innerLen -gt 0) {
                $innerTrim = $t.Substring($openParen + 1, $innerLen).Trim()
                if ($innerTrim -match '(?is)^\s*SELECT\b') {
                    $subRows = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $innerTrim -OperationType "SELECT" -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)
                    $roughFromIn = Get-OracleRoughFromClauseAfterFromKeyword -SelectSql $innerTrim
                    $amInInner = Get-OracleFromAliasToTableMap -FromClause $roughFromIn
                    foreach ($sr in $subRows) {
                        if ($null -eq $sr -or $null -eq $sr.TableName -or $sr.TableName -eq '') { continue }
                        $tnIn = [string]$sr.TableName
                        if ($amInInner.ContainsKey($tnIn)) {
                            $tnIn = [string]$amInInner[$tnIn]
                        }
                        $dup = $false
                        foreach ($existing in $out) {
                            if ($null -ne $existing.TableName -and $existing.TableName -eq $tnIn -and $existing.ColumnName -eq $sr.ColumnName -and $existing.Operation -eq $sr.Operation) {
                                $dup = $true
                                break
                            }
                        }
                        if (-not $dup) {
                            [void]$out.Add(@{
                                TableName  = $tnIn
                                ColumnName = $sr.ColumnName
                                Operation  = $sr.Operation
                            })
                        }
                    }
                }
            }
        }
        $searchPos = $j
    }

    return ,[object[]]@($out.ToArray())
}

function Split-ByCommaRespectingParens {
    param([string]$Text)

    $parts = [System.Collections.ArrayList]::new()
    $depth = 0
    $inString = $false
    $current = [System.Text.StringBuilder]::new()
    $len = $Text.Length
    $i = 0
    while ($i -lt $len) {
        $ch = $Text[$i]
        if ($inString) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $len -and $Text[$i + 1] -eq [char]0x27) {
                [void]$current.Append($ch)
                [void]$current.Append($Text[$i + 1])
                $i += 2
                continue
            }
            [void]$current.Append($ch)
            if ($ch -eq [char]0x27) {
                $inString = $false
            }
            $i++
            continue
        }
        if ($ch -eq [char]0x27) {
            $inString = $true
            [void]$current.Append($ch)
            $i++
            continue
        }
        if ($ch -eq '(') {
            $depth++
        }
        elseif ($ch -eq ')') {
            $depth--
        }

        if ($ch -eq ',' -and $depth -eq 0) {
            [void]$parts.Add($current.ToString())
            $current = [System.Text.StringBuilder]::new()
        }
        else {
            [void]$current.Append($ch)
        }
        $i++
    }
    if ($current.ToString().Trim() -ne '') {
        [void]$parts.Add($current.ToString())
    }

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

function Normalize-OracleTableList {
    param($Raw)

    if ($null -eq $Raw) {
        return ,[string[]]@()
    }
    if ($Raw -is [string]) {
        return ,[string[]]@([string]$Raw)
    }
    # ArrayList と、関数戻り値で PowerShell が展開した Object[] の両方を扱う（後者を [string] キャストすると空白区切り1文字列になる）
    if ($Raw -is [System.Collections.IList]) {
        if ($Raw.Count -eq 0) {
            return ,[string[]]@()
        }
        $arr = [string[]]::new($Raw.Count)
        for ($i = 0; $i -lt $Raw.Count; $i++) {
            $arr[$i] = [string]$Raw[$i]
        }
        return ,$arr
    }
    return ,[string[]]@([string]$Raw)
}

function Normalize-CrudRowList {
    param($Raw)

    if ($null -eq $Raw) {
        return ,[object[]]@()
    }
    if ($Raw -is [hashtable]) {
        return ,[object[]]@($Raw)
    }
    if ($Raw -is [System.Collections.ArrayList]) {
        if ($Raw.Count -eq 0) {
            return ,[object[]]@()
        }
        $arr = [object[]]::new($Raw.Count)
        for ($i = 0; $i -lt $Raw.Count; $i++) {
            $arr[$i] = $Raw[$i]
        }
        return ,$arr
    }
    if ($Raw -is [object[]]) {
        return ,$Raw
    }
    return ,[object[]]@($Raw)
}

function Test-OraclePlSqlDynamicSqlVarTableName {
    param([string]$Name)
    if ($null -eq $Name -or $Name -eq '') {
        return $false
    }
    return ($Name.ToUpper() -eq 'V_SQL')
}

function Get-FromTables {
    param(
        [string]$FromClause,
        [System.Collections.Generic.HashSet[string]]$ExcludeNames = $null
    )

    $tables = [System.Collections.ArrayList]::new()

    $cleaned = $FromClause -replace '(?i)\b(LEFT|RIGHT|FULL)\s+OUTER\s+JOIN\b', ','
    $cleaned = $cleaned -replace '(?i)\b(INNER|LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL)\s+JOIN\b', ','
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
        if ($trimmed -match '(?is)^TABLE\s*\(') {
            continue
        }
        if ($trimmed -match '(?is)^v_sql\s+\S') {
            continue
        }

        $trimmed = $trimmed -replace '(?i)\bWHERE\b.*$', ''
        $trimmed = $trimmed -replace '(?i)\bON\b.*$', ''
        $trimmed = $trimmed.Trim()
        if ($trimmed -eq '') { continue }
        $trimmed = $trimmed -replace '(?i)^(LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL|INNER)\s+', ''
        $trimmed = $trimmed -replace '(?i)\s+AS\s+', ' '

        if ($trimmed -match '(?:([\w$]+)\.)?([\w$]+)(?:\s+([\w$]+))?') {
            $firstId = [string]$Matches[1]
            if ($firstId -and $firstId -match '^(?i)GP_[TV]\d+$') {
                continue
            }
            $tblName = $Matches[2].ToUpper()
            if (Test-OraclePlSqlDynamicSqlVarTableName -Name $tblName) {
                continue
            }
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

function Get-OracleFromAliasToTableMap {
    param([string]$FromClause)

    $map = [System.Collections.Generic.Dictionary[string, string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if ($null -eq $FromClause -or $FromClause.Trim() -eq '') {
        return $map
    }

    $cleaned = $FromClause -replace '(?i)\b(LEFT|RIGHT|FULL)\s+OUTER\s+JOIN\b', ','
    $cleaned = $cleaned -replace '(?i)\b(INNER|LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL)\s+JOIN\b', ','
    $cleaned = $cleaned -replace '(?i)\bJOIN\b', ','
    $parts = Split-ByCommaRespectingParens -Text $cleaned

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }
        if ($trimmed.StartsWith('(')) { continue }
        if ($trimmed -match '(?is)^TABLE\s*\(') {
            continue
        }
        if ($trimmed -match '(?is)^v_sql\s+\S') {
            continue
        }

        $trimmed = $trimmed -replace '(?i)\bWHERE\b.*$', ''
        $trimmed = $trimmed -replace '(?i)\bON\b.*$', ''
        $trimmed = $trimmed.Trim()
        if ($trimmed -eq '') { continue }
        $trimmed = $trimmed -replace '(?i)^(LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL|INNER)\s+', ''
        $trimmed = $trimmed -replace '(?i)\s+AS\s+', ' '

        if ($trimmed -match '(?:([\w$]+)\.)?([\w$]+)(?:\s+([\w$]+))?') {
            $firstId = [string]$Matches[1]
            if ($firstId -and $firstId -match '^(?i)GP_[TV]\d+$') {
                continue
            }
            $tblName = $Matches[2].ToUpper()
            $aliasTok = [string]$Matches[3]
            if ($aliasTok -and $aliasTok.Trim() -ne '') {
                $aliasU = $aliasTok.Trim().ToUpper()
                if ($aliasU -eq 'AS') {
                    continue
                }
                if ($map.ContainsKey($aliasU)) {
                    if ($map[$aliasU] -eq $aliasU -and $tblName -ne $aliasU) {
                        $map[$aliasU] = $tblName
                    }
                }
                else {
                    [void]$map.Add($aliasU, $tblName)
                }
            }
        }
    }

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }
        if ($trimmed.StartsWith('(')) { continue }
        if ($trimmed -match '(?is)^TABLE\s*\(') {
            continue
        }
        if ($trimmed -match '(?is)^v_sql\s+\S') {
            continue
        }

        $trimmed = $trimmed -replace '(?i)\bWHERE\b.*$', ''
        $trimmed = $trimmed -replace '(?i)\bON\b.*$', ''
        $trimmed = $trimmed.Trim()
        if ($trimmed -eq '') { continue }
        $trimmed = $trimmed -replace '(?i)^(LEFT|RIGHT|FULL|CROSS|OUTER|NATURAL|INNER)\s+', ''
        $trimmed = $trimmed -replace '(?i)\s+AS\s+', ' '

        if ($trimmed -match '(?:([\w$]+)\.)?([\w$]+)(?:\s+([\w$]+))?') {
            $firstId = [string]$Matches[1]
            if ($firstId -and $firstId -match '^(?i)GP_[TV]\d+$') {
                continue
            }
            $tblName = $Matches[2].ToUpper()
            $aliasTok = [string]$Matches[3]
            $hasAlias = ($aliasTok -and $aliasTok.Trim() -ne '' -and $aliasTok.Trim().ToUpper() -ne 'AS')
            if (-not $hasAlias) {
                if (-not $map.ContainsKey($tblName)) {
                    [void]$map.Add($tblName, $tblName)
                }
            }
            else {
                if (-not $map.ContainsKey($tblName)) {
                    [void]$map.Add($tblName, $tblName)
                }
            }
        }
    }

    return $map
}

function Resolve-OraclePhysicalTableNamesOrdered {
    param(
        [object[]]$Tables,
        [System.Collections.Generic.Dictionary[string, string]]$AliasMap
    )

    $out = [System.Collections.ArrayList]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($t in $Tables) {
        if ($null -eq $t -or [string]::IsNullOrWhiteSpace([string]$t)) { continue }
        $tn = [string]$t
        $phys = $tn
        if ($null -ne $AliasMap -and $AliasMap.ContainsKey($tn)) {
            $phys = $AliasMap[$tn]
        }
        if (-not $seen.Contains($phys)) {
            [void]$seen.Add($phys)
            [void]$out.Add($phys)
        }
    }
    return ,[string[]]@($out.ToArray())
}

function Get-FromTablesFromNestedSelectsInSelectClause {
    param(
        [string]$SelectClause,
        [System.Collections.Generic.HashSet[string]]$ExcludeNames,
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames = $null
    )

    $out = [System.Collections.ArrayList]::new()
    if ($null -eq $SelectClause -or $SelectClause.Trim() -eq '') {
        return ,[string[]]@()
    }

    $len = $SelectClause.Length
    $i = 0
    while ($i -lt $len) {
        if ($SelectClause[$i] -eq '(' -and ($i + 1 -lt $len)) {
            $sub = $SelectClause.Substring($i)
            if ($sub -match '(?is)^\(\s*SELECT\b') {
                $depth = 0
                $inStr = $false
                $j = $i
                while ($j -lt $len) {
                    $adv = Step-OracleSqlScanOneChar -Text $SelectClause -ScanPos $j -InString ([ref]$inStr) -Depth ([ref]$depth)
                    $j += $adv
                    if ($depth -eq 0 -and $j -gt $i) { break }
                }
                if ($depth -eq 0 -and $j -gt $i + 1) {
                    $innerLen = $j - $i - 2
                    if ($innerLen -gt 0) {
                        $inner = $SelectClause.Substring($i + 1, $innerLen).Trim()
                        if ($inner -match '(?is)^SELECT\b') {
                            $nestedRows = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $inner -OperationType "SELECT" -AdditionalCteNames @($ExcludeNames) -PlSqlDeclaredNames $PlSqlDeclaredNames)
                            foreach ($nr in $nestedRows) {
                                if ($nr.Operation -eq 'R' -and $null -ne $nr.TableName -and $nr.TableName -ne '') {
                                    $dup = $false
                                    foreach ($o in $out) {
                                        if ($o -eq $nr.TableName) { $dup = $true; break }
                                    }
                                    if (-not $dup) { [void]$out.Add($nr.TableName) }
                                }
                            }
                        }
                    }
                }
                $i = $j
                continue
            }
        }
        $i++
    }

    return ,[string[]]@($out.ToArray())
}

function Get-SelectColumns {
    param([string]$SelectClause)

    $columns = [System.Collections.ArrayList]::new()

    if ($SelectClause.Trim() -eq '*') {
        [void]$columns.Add("*")
        return $columns
    }

    $cleaned = $SelectClause -replace '(?i)\bBULK\s+COLLECT\s+', ' '
    $cleaned = $cleaned -replace '(?i)\bDISTINCT\b', ''
    $cleaned = $cleaned -replace '(?is)\bINTO\b.*$', ''
    $cleaned = $cleaned.Trim()
    $parts = Split-ByCommaRespectingParens -Text $cleaned

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }

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

        $colExprAgg = $colExpr -replace '[\r\n]+', ' '
        if ($colExprAgg -match '(?i)(?<![\w$#])(?:COUNT|SUM|AVG|MIN|MAX)\s*\(\s*(?:DISTINCT\s+)?([^)]*)\)') {
            $innerAgg = $Matches[1].Trim() -replace '(?i)^DISTINCT\s+', ''
            if ($innerAgg -eq '') {
                continue
            }
            if ($innerAgg -eq '*' -or $innerAgg -match '^\*+$') {
                continue
            }
            if ($innerAgg -match '^\d+$') {
                continue
            }
            if ($innerAgg -match '^(?:[\w$]+\.)*([\w$]+)\s*$') {
                $innerName = $Matches[1]
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

function Get-SelectColumnRefs {
    param([string]$SelectClause)

    $refs = [System.Collections.ArrayList]::new()
    $starOnly = $false

    if ($SelectClause.Trim() -eq '*') {
        $starOnly = $true
        return @{ StarOnly = $starOnly; Refs = $refs }
    }

    $cleaned = $SelectClause -replace '(?i)\bBULK\s+COLLECT\s+', ' '
    $cleaned = $cleaned -replace '(?i)\bDISTINCT\b', ''
    $cleaned = $cleaned -replace '(?is)\bINTO\b.*$', ''
    $cleaned = $cleaned.Trim()
    $parts = Split-ByCommaRespectingParens -Text $cleaned

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        if ($trimmed -eq '') { continue }

        if ($trimmed -match '(?i)\bAS\s+(\w+)\s*$') {
            $colExpr = ($trimmed -replace '(?i)\s+AS\s+\w+\s*$', '').Trim()
        }
        elseif ($trimmed -match '\s+(\w+)\s*$') {
            $candidate = $Matches[1].ToUpper()
            $exprPart = ($trimmed -replace '\s+\w+\s*$', '').Trim()
            $isReserved = $candidate -in @('SELECT','FROM','WHERE','AND','OR','NOT','NULL','CASE','WHEN','THEN','ELSE','END','BY','ASC','DESC','INTO','AS')
            if ($exprPart -ne '' -and $candidate -notmatch '^\d' -and -not $isReserved -and -not (Test-SqlFunction -Name $candidate)) {
                $colExpr = $exprPart
            }
            else {
                $colExpr = $trimmed
            }
        }
        else {
            $colExpr = $trimmed
        }

        $colExprAgg = $colExpr -replace '[\r\n]+', ' '
        if ($colExprAgg -match '(?i)\bCASE\b') {
            foreach ($pr in (Get-ColumnRefsFromPredicateText -Text $colExprAgg)) {
                [void]$refs.Add($pr)
            }
            continue
        }
        if ($colExprAgg -match '(?i)(?<![\w$#])(?:COUNT|SUM|AVG|MIN|MAX)\s*\(\s*(?:DISTINCT\s+)?([^)]*)\)') {
            $innerAgg = $Matches[1].Trim() -replace '(?i)^DISTINCT\s+', ''
            if ($innerAgg -eq '') {
                continue
            }
            if ($innerAgg -eq '*' -or $innerAgg -match '^\*+$') {
                continue
            }
            if ($innerAgg -match '^\d+$') {
                continue
            }
            if ($innerAgg -match '^(?i)([\w$]+)\.([\w$]+)\.([\w$]+)\s*$') {
                $tblQ = $Matches[2].ToUpper()
                $innerName = $Matches[3].ToUpper()
                $isReservedCol = $innerName -in @('NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END')
                if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol) {
                    [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $innerName })
                }
                continue
            }
            if ($innerAgg -match '^(?i)([\w$]+)\.([\w$]+)\s*$') {
                $tblQ = $Matches[1].ToUpper()
                $innerName = $Matches[2].ToUpper()
                $isReservedCol = $innerName -in @('NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END')
                if (-not (Test-SqlFunction -Name $tblQ) -and -not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol) {
                    [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $innerName })
                }
                continue
            }
            if ($innerAgg -match '^(?i)([\w$]+)\s*$') {
                $innerName = $Matches[1].ToUpper()
                $isReservedCol = $innerName -in @('NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END')
                if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol) {
                    [void]$refs.Add(@{ TableName = $null; ColumnName = $innerName })
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
                $starOnly = $true
            }
            elseif ($funcContent -match '^(?i)([\w$]+)\.([\w$]+)\.([\w$]+)\s*$') {
                $innerName = $Matches[3].ToUpper()
                $tblQ = $Matches[2].ToUpper()
                $isReservedCol = $innerName -in @('NULL','CASE','WHEN','THEN','ELSE','END')
                if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol -and $innerName -notmatch '^\d') {
                    [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $innerName })
                }
            }
            elseif ($funcContent -match '^(?i)([\w$]+)\.([\w$]+)\s*$') {
                $tblQ = $Matches[1].ToUpper()
                $innerName = $Matches[2].ToUpper()
                $isReservedCol = $innerName -in @('NULL','CASE','WHEN','THEN','ELSE','END')
                if (-not (Test-SqlFunction -Name $tblQ) -and -not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol -and $innerName -notmatch '^\d') {
                    [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $innerName })
                }
            }
            elseif ($funcContent -match '(?i)^([\w$]+)$') {
                $innerName = $Matches[1].ToUpper()
                $isReservedCol = $innerName -in @('NULL','CASE','WHEN','THEN','ELSE','END')
                if (-not (Test-SqlFunction -Name $innerName) -and -not $isReservedCol -and $innerName -notmatch '^\d') {
                    [void]$refs.Add(@{ TableName = $null; ColumnName = $innerName })
                }
            }
            # スカラサブクエリ (SELECT ...) 内の WHERE / ON 等の列参照（例: WHERE C.ID = B.CAT_ID）
            if ($colExpr -match '(?is)\(\s*SELECT\b') {
                Bump-CrudParseProfile -Name 'GetSelectColumnRefsScalarExtra'
                foreach ($pr in (Get-ColumnRefsFromPredicateText -Text $colExpr)) {
                    [void]$refs.Add($pr)
                }
            }
        }
        elseif ($colExpr -match '^(?i)([\w$]+)\.([\w$]+)\.([\w$]+)\s*$') {
            $colName = $Matches[3].ToUpper()
            $tblQ = $Matches[2].ToUpper()
            $isReservedCol2 = $colName -in @('SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'INTO')
            if (-not $isReservedCol2 -and $colName -notmatch '^\d+$' -and -not (Test-SqlFunction -Name $colName)) {
                [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $colName })
            }
        }
        elseif ($colExpr -match '^(?i)([\w$]+)\.([\w$]+)\s*$') {
            $tblQ = $Matches[1].ToUpper()
            $colName = $Matches[2].ToUpper()
            $isReservedCol2 = $colName -in @('SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'INTO')
            if (-not (Test-SqlFunction -Name $tblQ) -and -not $isReservedCol2 -and $colName -notmatch '^\d+$' -and -not (Test-SqlFunction -Name $colName)) {
                [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $colName })
            }
        }
        elseif ($colExpr -match '(?:\w+\.)?(\w+)$') {
            $colName = $Matches[1].ToUpper()
            $isReservedCol2 = $colName -in @('SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'INTO')
            if (-not $isReservedCol2 -and $colName -notmatch '^\d+$' -and -not (Test-SqlFunction -Name $colName)) {
                [void]$refs.Add(@{ TableName = $null; ColumnName = $colName })
            }
        }
    }

    if ($refs.Count -eq 0 -and $SelectClause -match '(?i)\(\s*\*\s*\)') {
        $mStar = [regex]::Match($SelectClause, '(?i)\(\s*\*\s*\)')
        if ($mStar.Success) {
            $beforeStar = $SelectClause.Substring(0, $mStar.Index).TrimEnd()
            if ($beforeStar -notmatch '(?i)(?:COUNT|SUM|AVG|MIN|MAX)\s*$') {
                $starOnly = $true
            }
        }
    }

    return @{ StarOnly = $starOnly; Refs = $refs }
}

function Get-SetColumns {
    param([string]$SetClause)

    $columns = [System.Collections.ArrayList]::new()
    $st = if ($null -eq $SetClause) { '' } else { $SetClause.Trim() }
    if ($st.Length -ge 2 -and $st[0] -eq '(') {
        $eqPosT = Get-FirstEqualsAtParenDepthZero -Text $st
        if ($eqPosT -gt 0) {
            $jx = $eqPosT - 1
            while ($jx -ge 0 -and [char]::IsWhiteSpace($st[$jx])) {
                $jx--
            }
            if ($jx -ge 0 -and $st[$jx] -eq ')') {
                $lhsInner = Get-OracleBalancedParenInner -Text $st -OpenIndex 0
                if ($null -ne $lhsInner -and $lhsInner.Trim() -ne '') {
                    foreach ($seg in (Split-ByCommaRespectingParens -Text $lhsInner)) {
                        $cell = $seg.Trim()
                        if ($cell -eq '') {
                            continue
                        }
                        if ($cell -match '(?i)^([\w$]+(?:\.[\w$]+)*)') {
                            $lhsSegs = $Matches[1].Split('.')
                            [void]$columns.Add($lhsSegs[-1].ToUpper())
                        }
                    }
                }
            }
        }
        if ($columns.Count -gt 0) {
            return $columns
        }
    }
    $parts = Split-ByCommaRespectingParens -Text $SetClause

    foreach ($part in $parts) {
        $trimmed = $part.Trim()
        # 別名付き UPDATE（例: SET U.COL1 = 'X'）では左辺が U.COL1 となるため、最後の識別子を列名とする
        if ($trimmed -match '^([\w$]+(?:\.[\w$]+)*)\s*=') {
            $lhs = [string]$Matches[1]
            $lhsSegs = $lhs.Split('.')
            [void]$columns.Add($lhsSegs[-1].ToUpper())
        }
    }

    return $columns
}

function Get-ColumnRefsFromPredicateText {
    param([string]$Text)

    Bump-CrudParseProfile -Name 'GetColumnRefsFromPredicateText'

    $refs = [System.Collections.ArrayList]::new()
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $refs
    }
    $t = $Text -replace "'[^']*(?:''[^']*)*'", ' '
    $kwLeft = @(
        'SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'AS', 'ON', 'IN', 'IS', 'LIKE',
        'BETWEEN', 'EXISTS', 'JOIN', 'LEFT', 'RIGHT', 'INNER', 'OUTER', 'CROSS', 'GROUP', 'ORDER', 'BY', 'HAVING', 'UNION', 'INTERSECT',
        'MINUS', 'FETCH', 'OFFSET', 'RETURNING', 'OVER', 'PARTITION', 'CONNECT', 'START', 'WITH', 'PRIOR', 'DUAL', 'ROWNUM', 'SYSDATE'
    )
    $kwCol = @(
        'SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT', 'NULL', 'CASE', 'WHEN', 'THEN', 'ELSE', 'END', 'ASC', 'DESC', 'TRUE', 'FALSE', 'ON'
    )
    $rx3 = [regex]'(?i)\b([\w$]+)\.([\w$]+)\.([\w$]+)\b'
    foreach ($m in $rx3.Matches($t)) {
        $tblQ = $m.Groups[2].Value.ToUpper()
        $colQ = $m.Groups[3].Value.ToUpper()
        if ($tblQ -in $kwLeft) { continue }
        if ($colQ -in $kwCol) { continue }
        if (Test-SqlFunction -Name $colQ) { continue }
        [void]$refs.Add(@{ TableName = $tblQ; ColumnName = $colQ })
    }
    $rx2 = [regex]'(?i)\b([\w$]+)\.([\w$]+)\b'
    foreach ($m in $rx2.Matches($t)) {
        $a = $m.Groups[1].Value.ToUpper()
        $b = $m.Groups[2].Value.ToUpper()
        if ($a -in $kwLeft) { continue }
        if ($b -in $kwCol) { continue }
        if (Test-SqlFunction -Name $a) { continue }
        if (Test-SqlFunction -Name $b) { continue }
        [void]$refs.Add(@{ TableName = $a; ColumnName = $b })
    }
    return $refs
}

function Get-OracleBalancedParenInner {
    param(
        [string]$Text,
        [int]$OpenIndex
    )
    if ($null -eq $Text -or $OpenIndex -lt 0 -or $OpenIndex -ge $Text.Length) {
        return $null
    }
    if ($Text[$OpenIndex] -ne '(') {
        return $null
    }
    $depth = 0
    $inStr = $false
    for ($i = $OpenIndex; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inStr) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq [char]0x27) {
                $i++
                continue
            }
            if ($ch -eq [char]0x27) {
                $inStr = $false
            }
            continue
        }
        if ($ch -eq [char]0x27) {
            $inStr = $true
            continue
        }
        if ($ch -eq '(') {
            $depth++
        }
        elseif ($ch -eq ')') {
            $depth--
            if ($depth -eq 0) {
                return $Text.Substring($OpenIndex + 1, $i - $OpenIndex - 1).Trim()
            }
        }
    }
    return $null
}

function Get-FirstEqualsAtParenDepthZero {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return -1
    }
    $depth = 0
    $inStr = $false
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $ch = $Text[$i]
        if ($inStr) {
            if ($ch -eq [char]0x27 -and $i + 1 -lt $Text.Length -and $Text[$i + 1] -eq [char]0x27) {
                $i++
                continue
            }
            if ($ch -eq [char]0x27) {
                $inStr = $false
            }
            continue
        }
        if ($ch -eq [char]0x27) {
            $inStr = $true
            continue
        }
        if ($ch -eq '(') {
            $depth++
        }
        elseif ($ch -eq ')') {
            $depth--
            if ($depth -lt 0) {
                return -1
            }
        }
        elseif ($ch -eq '=' -and $depth -eq 0) {
            return $i
        }
    }
    return -1
}

function Get-PhysicalTableNamesForMergeUsingSelect {
    param(
        [string]$InnerSelectSql,
        [System.Collections.Generic.HashSet[string]]$CteNames,
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames,
        [string[]]$AdditionalCteNames
    )

    if ([string]::IsNullOrWhiteSpace($InnerSelectSql)) {
        return ,[string[]]@()
    }
    $frag = $InnerSelectSql.Trim()
    if ($frag -notmatch '(?is)^SELECT\b') {
        return ,[string[]]@()
    }
    $rows = @(Get-TableAndColumns -SqlFragment $frag -OperationType 'SELECT' -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)
    $physNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $rows) {
        if ($r.Operation -ne 'R') {
            continue
        }
        $tn = [string]$r.TableName
        if ([string]::IsNullOrWhiteSpace($tn)) {
            continue
        }
        $tu = $tn.Trim().ToUpper()
        if ($tu -eq 'DUAL' -or $tu.EndsWith('.DUAL')) {
            continue
        }
        if ($null -ne $CteNames -and $CteNames.Count -gt 0 -and $CteNames.Contains($tu)) {
            continue
        }
        [void]$physNameSet.Add($tu)
    }
    $nameList = [System.Collections.ArrayList]::new()
    foreach ($pn in $physNameSet) {
        [void]$nameList.Add($pn)
    }
    return ,[string[]]@($nameList.ToArray())
}

function Add-UpdateSetRhsReadRows {
    param(
        [string]$SetClause,
        [hashtable]$OuterAliasToTable,
        [string]$DefaultTableName,
        [System.Collections.Generic.HashSet[string]]$CteNames,
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames,
        [string[]]$AdditionalCteNames,
        [System.Collections.ArrayList]$OutList
    )

    if ([string]::IsNullOrWhiteSpace($SetClause)) {
        return
    }
    $partsToProcess = [System.Collections.ArrayList]::new()
    $stAll = $SetClause.Trim()
    $tupleRhsDirect = $null
    if ($stAll.Length -ge 2 -and $stAll[0] -eq '(') {
        $depthT = 0
        $inStrT = $false
        for ($ti = 0; $ti -lt $stAll.Length; $ti++) {
            $cht = $stAll[$ti]
            if ($inStrT) {
                if ($cht -eq [char]0x27 -and $ti + 1 -lt $stAll.Length -and $stAll[$ti + 1] -eq [char]0x27) {
                    $ti++
                    continue
                }
                if ($cht -eq [char]0x27) {
                    $inStrT = $false
                }
                continue
            }
            if ($cht -eq [char]0x27) {
                $inStrT = $true
                continue
            }
            if ($cht -eq '(') {
                $depthT++
            }
            elseif ($cht -eq ')') {
                $depthT--
            }
            elseif ($cht -eq '=' -and $depthT -eq 0) {
                $tj = $ti - 1
                while ($tj -ge 0 -and [char]::IsWhiteSpace($stAll[$tj])) {
                    $tj--
                }
                if ($tj -ge 0 -and $stAll[$tj] -eq ')') {
                    $tupleRhsDirect = $stAll.Substring($ti + 1).Trim()
                    break
                }
            }
        }
    }
    if ($null -ne $tupleRhsDirect) {
        [void]$partsToProcess.Add($tupleRhsDirect)
    }
    else {
        foreach ($p0 in (Split-ByCommaRespectingParens -Text $SetClause)) {
            [void]$partsToProcess.Add($p0)
        }
    }
    foreach ($part in $partsToProcess) {
        $t = $part.Trim()
        if ($t -eq '') {
            continue
        }
        $rhs = $t
        if ($null -eq $tupleRhsDirect) {
            $eqPos = Get-FirstEqualsAtParenDepthZero -Text $t
            if ($eqPos -lt 0) {
                continue
            }
            $rhs = $t.Substring($eqPos + 1).Trim()
        }
        if ($rhs -eq '') {
            continue
        }
        $selFrag = $null
        if ($rhs.Length -ge 2 -and $rhs[0] -eq '(') {
            $inner = Get-OracleBalancedParenInner -Text $rhs -OpenIndex 0
            if ($null -ne $inner -and $inner.Trim() -match '(?is)^SELECT\b') {
                $selFrag = $inner.Trim()
            }
        }
        if ($null -ne $selFrag) {
            $roughInnerFrom = Get-OracleRoughFromClauseAfterFromKeyword -SelectSql $selFrag
            $innerAm = Get-OracleFromAliasToTableMap -FromClause $roughInnerFrom
            $innerRows = @(Get-TableAndColumns -SqlFragment $selFrag -OperationType 'SELECT' -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $PlSqlDeclaredNames)
            foreach ($ir in $innerRows) {
                if ($ir.Operation -ne 'R') {
                    continue
                }
                $tn = [string]$ir.TableName
                if ([string]::IsNullOrWhiteSpace($tn)) {
                    continue
                }
                $origQual = $tn.Trim().ToUpper()
                $phys = $origQual
                if ($null -ne $innerAm -and $innerAm.ContainsKey($origQual)) {
                    $phys = [string]$innerAm[$origQual]
                }
                elseif ($null -ne $OuterAliasToTable -and $OuterAliasToTable.ContainsKey($origQual)) {
                    $phys = [string]$OuterAliasToTable[$origQual]
                }
                else {
                    $phys = $tn.Trim()
                }
                $tu = $phys.Trim().ToUpper()
                if ($tu -eq 'DUAL' -or $tu.EndsWith('.DUAL')) {
                    continue
                }
                if ($null -ne $CteNames -and $CteNames.Count -gt 0 -and $CteNames.Contains($tu)) {
                    continue
                }
                if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $PlSqlDeclaredNames.Contains($tu)) {
                    continue
                }
                $cn = [string]$ir.ColumnName
                if ([string]::IsNullOrWhiteSpace($cn) -or $cn -eq '*' -or $cn -eq '(ALL)') {
                    continue
                }
                if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $PlSqlDeclaredNames.Contains($cn.ToUpper())) {
                    continue
                }
                [void]$OutList.Add(@{ TableName = $phys.Trim(); ColumnName = $cn.ToUpper(); Operation = 'R' })
            }
        }
        else {
            foreach ($r in (Get-ColumnRefsFromPredicateText -Text $rhs)) {
                $qual = if ($null -ne $r.TableName -and $r.TableName -ne '') { $r.TableName.ToUpper() } else { '' }
                if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                    continue
                }
                $cn = $r.ColumnName.ToUpper()
                if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $PlSqlDeclaredNames.Contains($cn)) {
                    continue
                }
                $phys = $null
                if ($qual -ne '') {
                    if ($null -ne $OuterAliasToTable -and $OuterAliasToTable.ContainsKey($qual)) {
                        $phys = [string]$OuterAliasToTable[$qual]
                    }
                    else {
                        $phys = $qual
                    }
                }
                else {
                    $phys = $DefaultTableName
                }
                $physU = ([string]$phys).Trim().ToUpper()
                if ($physU -eq 'DUAL' -or $physU.EndsWith('.DUAL')) {
                    continue
                }
                [void]$OutList.Add(@{ TableName = $phys; ColumnName = $cn; Operation = 'R' })
            }
        }
    }
}

function Get-MergeCrudRowsDetailed {
    param(
        [string]$SqlFragment,
        [System.Collections.Generic.HashSet[string]]$CteNames,
        [System.Collections.Generic.HashSet[string]]$PlSqlDeclaredNames = $null,
        [string[]]$AdditionalCteNames = @()
    )

    $out = [System.Collections.ArrayList]::new()
    $pattern = '(?i)MERGE\s+INTO\s+(?:([\w$]+)\.)?([\w$]+)(?:\s+(?:AS\s+)?([\w$]+))?\s+USING\s+'
    $allMatches = [regex]::Matches($SqlFragment, $pattern)
    $mi = 0
    foreach ($match in $allMatches) {
        $tgtTable = $match.Groups[2].Value.ToUpper()
        if ($CteNames.Count -gt 0 -and $CteNames.Contains($tgtTable)) { continue }
        $segStart = $match.Index
        $segEnd = $SqlFragment.Length
        if ($mi + 1 -lt $allMatches.Count) {
            $segEnd = $allMatches[$mi + 1].Index
        }
        $mergeSeg = $SqlFragment.Substring($segStart, $segEnd - $segStart)
        $mi++
        $thisOut = [System.Collections.ArrayList]::new()
        $tail = $mergeSeg.Substring($match.Length)
        $mergeUsingMultiPhysList = [string[]]@()
        $mergeUsingAliasTok = $null
        $mergeAliasToPhysical = @{}
        if (-not $mergeAliasToPhysical.ContainsKey($tgtTable)) {
            $mergeAliasToPhysical[$tgtTable] = $tgtTable
        }
        if ($match.Groups[3].Success -and $match.Groups[3].Value -ne '') {
            $tgtAliasTok = $match.Groups[3].Value.ToUpper()
            if ($tgtAliasTok -ne 'AS') {
                $mergeAliasToPhysical[$tgtAliasTok] = $tgtTable
            }
        }
        $mOn = [regex]::Match($tail, '(?is)\bON\s+')
        if ($mOn.Success -and $mOn.Index -gt 0) {
            $usingOnly = $tail.Substring(0, $mOn.Index).TrimEnd()
            $u = $usingOnly.TrimStart()
            if ($u.Length -gt 0 -and $u[0] -eq '(') {
                $depthParen = 0
                $inStrU = $false
                $closeAt = -1
                $ku = 0
                while ($ku -lt $u.Length) {
                    $chu = $u[$ku]
                    if ($inStrU) {
                        if ($chu -eq [char]0x27 -and $ku + 1 -lt $u.Length -and $u[$ku + 1] -eq [char]0x27) { $ku += 2; continue }
                        if ($chu -eq [char]0x27) { $inStrU = $false }
                        $ku++
                        continue
                    }
                    if ($chu -eq [char]0x27) { $inStrU = $true; $ku++; continue }
                    if ($chu -eq '(') { $depthParen++ }
                    elseif ($chu -eq ')') {
                        $depthParen--
                        if ($depthParen -eq 0) { $closeAt = $ku; break }
                    }
                    $ku++
                }
                if ($closeAt -gt 0) {
                    $innerSql = $u.Substring(1, $closeAt - 1).Trim()
                    $afterSub = $u.Substring($closeAt + 1).TrimStart()
                    $usingAliasName = $null
                    $mAfterAlias = [regex]::Match($afterSub, '(?is)^(?:AS\s+)?([\w$]+)\s*$')
                    if ($mAfterAlias.Success -and $mAfterAlias.Groups[1].Success -and $mAfterAlias.Groups[1].Value -ne '') {
                        $usingAliasName = $mAfterAlias.Groups[1].Value.ToUpper()
                    }
                    $mergeUsingAliasTok = $usingAliasName
                    $roughFrom = ''
                    $mFromKw = [regex]::Match($innerSql, '(?is)\bFROM\s+')
                    if ($mFromKw.Success) {
                        $fromRest = $innerSql.Substring($mFromKw.Index + $mFromKw.Length)
                        $cutLen = $fromRest.Length
                        foreach ($term in @('WHERE', 'GROUP', 'ORDER', 'HAVING', 'UNION', 'INTERSECT', 'MINUS', 'FETCH')) {
                            $tm = [regex]::Match($fromRest, "(?is)\b$term\b")
                            if ($tm.Success -and $tm.Index -lt $cutLen) { $cutLen = $tm.Index }
                        }
                        $roughFrom = $fromRest.Substring(0, [Math]::Min($cutLen, $fromRest.Length)).Trim()
                    }
                    if ($roughFrom -ne '') {
                        $amU = Get-OracleFromAliasToTableMap -FromClause $roughFrom
                        foreach ($dek in $amU.Keys) {
                            $mergeAliasToPhysical[$dek] = $amU[$dek]
                        }
                    }
                    $physFromInner = @(Get-PhysicalTableNamesForMergeUsingSelect -InnerSelectSql $innerSql -CteNames $CteNames -PlSqlDeclaredNames $PlSqlDeclaredNames -AdditionalCteNames $AdditionalCteNames)
                    if ($null -ne $usingAliasName -and $usingAliasName -ne '') {
                        if ($physFromInner.Count -eq 1) {
                            $mergeAliasToPhysical[$usingAliasName] = [string]$physFromInner[0]
                        }
                        elseif ($physFromInner.Count -gt 1) {
                            $mergeUsingMultiPhysList = [string[]]@($physFromInner)
                        }
                    }
                }
            }
            else {
                if ($u -match '(?is)^(?:([\w$]+)\.)?([\w$]+)(?:\s+(?:AS\s+)?([\w$]+))?\s*$') {
                    $usingPhys = ([string]$Matches[2]).ToUpper()
                    if ($CteNames.Count -eq 0 -or -not $CteNames.Contains($usingPhys)) {
                        if (-not $mergeAliasToPhysical.ContainsKey($usingPhys)) {
                            $mergeAliasToPhysical[$usingPhys] = $usingPhys
                        }
                        if ($Matches.ContainsKey(3) -and [string]$Matches[3] -ne '') {
                            $usal = ([string]$Matches[3]).ToUpper()
                            if ($usal -ne 'AS') {
                                $mergeAliasToPhysical[$usal] = $usingPhys
                            }
                        }
                    }
                }
            }
        }
        if ($mOn.Success) {
            $afterOnPos = $mOn.Index + $mOn.Length
            $whenM = [regex]::Match($tail, '(?is)\bWHEN\s+(?:NOT\s+)?MATCHED\b')
            $onEndPos = $tail.Length
            if ($whenM.Success -and $whenM.Index -gt $afterOnPos) {
                $onEndPos = $whenM.Index
            }
            $onBodyForRefs = $tail.Substring($afterOnPos, $onEndPos - $afterOnPos).Trim()
            if ($onBodyForRefs -ne '') {
                $onRefs = Get-ColumnRefsFromPredicateText -Text $onBodyForRefs
                foreach ($r in $onRefs) {
                    $qual = if ($null -ne $r.TableName -and $r.TableName -ne '') { $r.TableName.ToUpper() } else { '' }
                    if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $qual -ne '' -and $PlSqlDeclaredNames.Contains($qual)) {
                        continue
                    }
                    $cn = $r.ColumnName.ToUpper()
                    $handledMulti = $false
                    if ($qual -ne '' -and $null -ne $mergeUsingAliasTok -and $qual -eq $mergeUsingAliasTok -and $mergeUsingMultiPhysList.Count -gt 1) {
                        foreach ($mtp in $mergeUsingMultiPhysList) {
                            $physU2 = $mtp.Trim().ToUpper()
                            if ($physU2 -eq 'DUAL' -or $physU2.EndsWith('.DUAL')) {
                                continue
                            }
                            [void]$thisOut.Add(@{ TableName = $mtp.Trim(); ColumnName = $cn; Operation = 'R' })
                        }
                        $handledMulti = $true
                    }
                    if (-not $handledMulti) {
                        $physTbl = $qual
                        if ($qual -ne '' -and $mergeAliasToPhysical.ContainsKey($qual)) {
                            $physTbl = [string]$mergeAliasToPhysical[$qual]
                        }
                        $physU = ([string]$physTbl).Trim().ToUpper()
                        if ($physU -eq 'DUAL' -or $physU.EndsWith('.DUAL')) {
                            continue
                        }
                        [void]$thisOut.Add(@{ TableName = $physTbl; ColumnName = $cn; Operation = 'R' })
                    }
                }
            }
        }
        $rxUpd = [regex]::Match($mergeSeg, '(?is)\bWHEN\s+MATCHED\b\s+THEN\s+UPDATE\s+SET\s+(.+?)(?=\bWHEN\b|\z)')
        if ($rxUpd.Success) {
            $setCols = Get-SetColumns -SetClause $rxUpd.Groups[1].Value
            foreach ($c in $setCols) {
                [void]$thisOut.Add(@{ TableName = $tgtTable; ColumnName = $c; Operation = 'U' })
            }
            $mergeReadMap = @{}
            foreach ($mk in $mergeAliasToPhysical.Keys) {
                $mergeReadMap[$mk] = $mergeAliasToPhysical[$mk]
            }
            Add-UpdateSetRhsReadRows -SetClause $rxUpd.Groups[1].Value -OuterAliasToTable $mergeReadMap -DefaultTableName $tgtTable `
                -CteNames $CteNames -PlSqlDeclaredNames $PlSqlDeclaredNames -AdditionalCteNames $AdditionalCteNames -OutList $thisOut
        }
        $rxIns = [regex]::Match($mergeSeg, '(?is)\bWHEN\s+NOT\s+MATCHED\b\s+THEN\s+INSERT\s*\(([^)]+)\)')
        if ($rxIns.Success) {
            $insCols = ($rxIns.Groups[1].Value -split ',') | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne '' }
            foreach ($c in $insCols) {
                [void]$thisOut.Add(@{ TableName = $tgtTable; ColumnName = $c; Operation = 'C' })
            }
            $tailFromIns = $mergeSeg.Substring($rxIns.Index + $rxIns.Length)
            $mv = [regex]::Match($tailFromIns, '(?is)^\s*VALUES\s*\(')
            if ($mv.Success) {
                $absOpen = $rxIns.Index + $rxIns.Length + $mv.Index + $mv.Length - 1
                $valsInner = Get-OracleBalancedParenInner -Text $mergeSeg -OpenIndex $absOpen
                if ($null -ne $valsInner -and $valsInner.Trim() -ne '') {
                    $valRefs = Get-ColumnRefsFromPredicateText -Text $valsInner
                    foreach ($vr in $valRefs) {
                        $vqual = if ($null -ne $vr.TableName -and $vr.TableName -ne '') { $vr.TableName.ToUpper() } else { '' }
                        if ($null -ne $PlSqlDeclaredNames -and $PlSqlDeclaredNames.Count -gt 0 -and $vqual -ne '' -and $PlSqlDeclaredNames.Contains($vqual)) {
                            continue
                        }
                        $vcn = $vr.ColumnName.ToUpper()
                        $vhandled = $false
                        if ($vqual -ne '' -and $null -ne $mergeUsingAliasTok -and $vqual -eq $mergeUsingAliasTok -and $mergeUsingMultiPhysList.Count -gt 1) {
                            foreach ($mtp in $mergeUsingMultiPhysList) {
                                $physU3 = $mtp.Trim().ToUpper()
                                if ($physU3 -eq 'DUAL' -or $physU3.EndsWith('.DUAL')) {
                                    continue
                                }
                                [void]$thisOut.Add(@{ TableName = $mtp.Trim(); ColumnName = $vcn; Operation = 'R' })
                            }
                            $vhandled = $true
                        }
                        if (-not $vhandled) {
                            $vphys = $vqual
                            if ($vqual -ne '' -and $mergeAliasToPhysical.ContainsKey($vqual)) {
                                $vphys = [string]$mergeAliasToPhysical[$vqual]
                            }
                            $vphysU = ([string]$vphys).Trim().ToUpper()
                            if ($vphysU -eq 'DUAL' -or $vphysU.EndsWith('.DUAL')) {
                                continue
                            }
                            [void]$thisOut.Add(@{ TableName = $vphys; ColumnName = $vcn; Operation = 'R' })
                        }
                    }
                }
            }
        }
        if ($thisOut.Count -eq 0) {
            [void]$thisOut.Add(@{ TableName = $tgtTable; ColumnName = '(ALL)'; Operation = 'C' })
            [void]$thisOut.Add(@{ TableName = $tgtTable; ColumnName = '(ALL)'; Operation = 'U' })
        }
        foreach ($row in $thisOut) { [void]$out.Add($row) }
    }
    return $out
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

function Get-OracleExecuteImmediateLiteralSqlFragments {
    param([string]$Content)

    $out = [System.Collections.ArrayList]::new()
    if ([string]::IsNullOrEmpty($Content)) {
        return $out
    }
    # マスク前の本文から EXECUTE IMMEDIATE '...' 内の SQL を取り出す（Python extract_dynamic_sql_literals に相当）
    $pattern = "(?is)\bEXECUTE\s+IMMEDIATE\s+'((?:[^']|'')*)'"
    foreach ($m in [regex]::Matches($Content, $pattern)) {
        $inner = $m.Groups[1].Value -replace "''", "'"
        if ($inner -match '(?i)\b(SELECT|INSERT|UPDATE|DELETE|MERGE)\b') {
            [void]$out.Add($inner)
        }
    }
    return $out
}

function ConvertFrom-OracleSqlFile {
    param(
        [string]$FilePath,
        [string[]]$AdditionalCteNames = @(),
        [switch]$DebugLog
    )

    $rawContent = Get-Content $FilePath -Raw -Encoding Default
    $contentNoComments = Remove-SqlComments -Content $rawContent
    $content = Mask-OracleSqlStringLiteralsForParse -Content $contentNoComments
    $fileName = [System.IO.Path]::GetFileName($FilePath)

    $objectInfo = Get-OracleObjectInfo -Content $content -FileName $fileName

    $parseContent = $content
    if ($objectInfo.ObjectType -eq "PACKAGE") {
        $parseContent = Get-PackageBodySection -Content $content
    }

    $dynamicSource = $contentNoComments
    if ($objectInfo.ObjectType -eq "PACKAGE") {
        $dynamicSource = Get-PackageBodySection -Content $contentNoComments
    }
    $dynamicSqlFragments = Get-OracleExecuteImmediateLiteralSqlFragments -Content $dynamicSource

    $plsqlDeclNames = Get-OraclePlSqlDeclaredVariableNames -PlSqlBlock $parseContent

    Reset-CrudParseProfileStats

    $results = [System.Collections.ArrayList]::new()
    $featureName = "$($objectInfo.ObjectType):$($objectInfo.ObjectName)"
    $extractCount = 0

    foreach ($opType in @("INSERT", "SELECT", "UPDATE", "DELETE", "MERGE")) {
        try {
            $extracted = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $parseContent -OperationType $opType -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $plsqlDeclNames)
        }
        catch {
            $snippet = if ($parseContent.Length -gt 100) { $parseContent.Substring(0, 100) + "..." } else { $parseContent }
            Write-Warning "[Oracle] 解析エラー詳細: $fileName | 操作=$opType | $($_.Exception.Message)"
            Write-Warning "[Oracle] SQL断片(先頭): $snippet"
            Write-Warning "[Oracle] スタック: $($_.ScriptStackTrace)"
            $extracted = @()
        }
        $extractCount += $extracted.Count

        foreach ($item in $extracted) {
            [void]$results.Add(@{
                SourceType  = "Oracle"
                SourceFile  = $fileName
                ObjectType  = $objectInfo.ObjectType
                ObjectName  = $objectInfo.ObjectName
                ProcName    = $objectInfo.ObjectName
                FeatureName = $featureName
                TableName   = $item.TableName
                ColumnName  = $item.ColumnName
                Operation   = $item.Operation
            })
        }
    }

    foreach ($dynSql in $dynamicSqlFragments) {
        foreach ($opType in @("INSERT", "SELECT", "UPDATE", "DELETE", "MERGE")) {
            try {
                $extracted = Normalize-CrudRowList (Get-TableAndColumns -SqlFragment $dynSql -OperationType $opType -AdditionalCteNames $AdditionalCteNames -PlSqlDeclaredNames $plsqlDeclNames)
            }
            catch {
                $snippet = if ($dynSql.Length -gt 100) { $dynSql.Substring(0, 100) + "..." } else { $dynSql }
                Write-Warning "[Oracle] 解析エラー詳細: $fileName | 操作=$opType (EXECUTE IMMEDIATE) | $($_.Exception.Message)"
                Write-Warning "[Oracle] SQL断片(先頭): $snippet"
                Write-Warning "[Oracle] スタック: $($_.ScriptStackTrace)"
                $extracted = @()
            }
            $extractCount += $extracted.Count

            foreach ($item in $extracted) {
                [void]$results.Add(@{
                    SourceType  = "Oracle"
                    SourceFile  = $fileName
                    ObjectType  = $objectInfo.ObjectType
                    ObjectName  = $objectInfo.ObjectName
                    ProcName    = $objectInfo.ObjectName
                    FeatureName = $featureName
                    TableName   = $item.TableName
                    ColumnName  = $item.ColumnName
                    Operation   = $item.Operation
                })
            }
        }
    }

    if ($DebugLog) {
        $hint = if ($extractCount -eq 0) { " (パーサで0件→SQL未検出の可能性)" } else { "" }
        Write-Host "[Oracle][Debug] 解析: $fileName | $($objectInfo.ObjectName) | 抽出=$extractCount | 本文=$($parseContent.Length) 文字$hint" -ForegroundColor DarkCyan
    }

    Write-CrudParseProfileReport -Label "$fileName | $($objectInfo.ObjectName) | $($parseContent.Length) chars"

    return $results
}

function ConvertFrom-OracleSqlDirectory {
    param(
        [string]$SourcePath,
        [string]$FilePattern = "*.sql",
        [string[]]$ExcludePatterns = @(),
        [string[]]$ExcludeTables = @(),
        [string[]]$ExcludeSchemas = @(),
        [string[]]$AdditionalCteNames = @(),
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
            $fileResults = ConvertFrom-OracleSqlFile -FilePath $file.FullName -AdditionalCteNames $AdditionalCteNames -DebugLog:$DebugLog

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
                            if ($r.TableName -match '\.') {
                                $schemaPrefix = ($r.TableName.Split('.')[0]).ToUpper()
                                foreach ($schema in $ExcludeSchemas) {
                                    if ([string]::IsNullOrWhiteSpace($schema)) { continue }
                                    if ($schemaPrefix -eq $schema.Trim().ToUpper()) {
                                        $skip = $true
                                        $reason = "ExcludeSchemas:$schema"
                                        break
                                    }
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
                if (-not $skip -and $result.TableName -match '\.') {
                    $schemaPrefix = ($result.TableName.Split('.')[0]).ToUpper()
                    foreach ($schema in $ExcludeSchemas) {
                        if ([string]::IsNullOrWhiteSpace($schema)) { continue }
                        if ($schemaPrefix -eq $schema.Trim().ToUpper()) { $skip = $true; break }
                    }
                }
                if (-not $skip) {
                    [void]$allResults.Add($result)
                }
            }
        }
        catch {
            Write-Warning "[Oracle] 解析エラー: $($file.FullName) - $($_.Exception.Message)"
            Write-Warning "[Oracle] スタック: $($_.ScriptStackTrace)"
        }
    }

    Write-Progress -Activity "Oracle SQL 解析中" -Completed
    Write-Host "[Oracle] 解析完了: $($allResults.Count) 件のCRUDエントリを検出" -ForegroundColor Green

    return $allResults
}
