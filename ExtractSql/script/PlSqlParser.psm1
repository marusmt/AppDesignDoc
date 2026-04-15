#Requires -Version 5.1
<#
.SYNOPSIS
    PL/SQLソースコードからSQL文を抽出するパーサーモジュール
.DESCRIPTION
    静的SQL、動的SQL（EXECUTE IMMEDIATE, DBMS_SQL, OPEN FOR）、
    IF分岐内のSQL断片をすべて抽出・展開します。
#>

using module .\SqlFormatter.psm1

# ============================================================
# 状態管理用Enum
# ============================================================
enum PlSqlParserState {
    Normal
    InStaticSql
    InDynamicSqlAssign
    InExecuteImmediate
    InDbmsSqlParse
    InOpenFor
    InIfBlock
}

# ============================================================
# Invoke-PlSqlParser: PL/SQLパース実行
# ============================================================
function Invoke-PlSqlParser {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string]$Encoding = 'UTF8',

        [Parameter()]
        [string]$LogFile = ''
    )

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    Write-Log -Level INFO -Message "Processing: $fileName (PL/SQL)" -LogFile $LogFile

    $lines = Get-Content -Path $FilePath -Encoding $Encoding
    $sqlStatements = [System.Collections.Generic.List[SqlStatement]]::new()

    $state = [PlSqlParserState]::Normal
    $currentSql = ''
    $startLine = 0
    $inBlockComment = $false
    $dynamicSqlVars = @{}  # 変数名 → @{Fragments; StartLine} のマッピング

    # 最後に更新された動的SQL変数の Fragments リストへの参照。
    # IF分岐の断片をその変数に直接追加するために使用する。
    $lastFragmentsList = $null

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lineNum = $i + 1
        $line = $lines[$i]
        $trimmed = $line.Trim()

        # ================================================
        # ブロックコメントのスキップ
        # ================================================
        if ($inBlockComment) {
            if ($trimmed -match '\*/') {
                $inBlockComment = $false
            }
            continue
        }
        if ($trimmed -match '^/\*' -and $trimmed -notmatch '\*/') {
            $inBlockComment = $true
            continue
        }

        # 行コメントのスキップ
        if ($trimmed -match '^--') {
            continue
        }

        # ================================================
        # IF分岐の検出と展開
        # ================================================
        if ($trimmed -match '(?i)^\s*IF\s+(.+?)\s+THEN') {
            # IF ブロック全体を収集
            $ifLines = [System.Collections.Generic.List[string]]::new()
            $ifStartLine = $lineNum
            $ifNestLevel = 0

            for ($j = $i; $j -lt $lines.Count; $j++) {
                $ifLine = $lines[$j].Trim()
                $ifLines.Add($lines[$j])

                if ($ifLine -match '(?i)^\s*IF\s+') {
                    $ifNestLevel++
                }
                if ($ifLine -match '(?i)^\s*END\s+IF') {
                    $ifNestLevel--
                    if ($ifNestLevel -eq 0) {
                        break
                    }
                }
            }

            # SQL断片抽出用のスクリプトブロック
            $extractSqlFromLine = {
                param($ln)
                $t = $ln.Trim()

                # 変数代入: v_sql := v_sql || '...'; or v_sql := '...';
                if ($t -match "(?i)^\w+\s*:=\s*(.+?)\s*;\s*$") {
                    $assignPart = $Matches[1]
                    return (Extract-PlSqlStringLiterals -Expression $assignPart)
                }

                # Append系: v_sql := v_sql || '...';
                if ($t -match "(?i)^(\w+)\s*:=\s*\1\s*\|\|\s*(.+?)\s*;\s*$") {
                    $appendPart = $Matches[2]
                    return (Extract-PlSqlStringLiterals -Expression $appendPart)
                }

                return $null
            }

            $branchResults = Expand-IfBranches -Lines $ifLines.ToArray() `
                -Language 'plsql' -ExtractSqlFromLine $extractSqlFromLine

            if ($branchResults.Count -gt 0) {
                # 直前に操作した動的SQL変数のFragmentsリストに直接追加する。
                # 変数が未確定の場合は後続の変数確定フェーズで拾われるようにフォールバック。
                if ($lastFragmentsList) {
                    foreach ($fragment in $branchResults) {
                        $lastFragmentsList.Add($fragment)
                    }
                }
                # 変数が特定できない場合は破棄（ログ警告を出す）
                else {
                    Write-Log -Level WARN -Message "Line ${ifStartLine}: IF分岐の断片を関連付ける動的SQL変数が見つかりません" -LogFile $LogFile
                }
            }

            # IFブロック分だけインデックスを進める
            $i = $j
            continue
        }

        # ================================================
        # EXECUTE IMMEDIATE の検出
        # ================================================
        if ($trimmed -match '(?i)^EXECUTE\s+IMMEDIATE\s+(.+)') {
            $execPart = $Matches[1]
            $startLine = $lineNum

            # 1行で完結する場合
            if ($execPart -match ';\s*$') {
                $execPart = $execPart -replace ';\s*$', ''
                $sql = Extract-PlSqlDynamicSql -Expression $execPart
                if ($sql) {
                    $stmt = [SqlStatement]::new()
                    $stmt.Sql = $sql
                    $stmt.Type = Get-SqlType -SqlText $sql
                    $stmt.Category = 'Dynamic'
                    $stmt.StartLine = $startLine
                    $stmt.EndLine = $lineNum
                    $stmt.SourceFile = $fileName
                    $sqlStatements.Add($stmt)
                }
            }
            else {
                # 複数行にまたがる場合
                $state = [PlSqlParserState]::InExecuteImmediate
                $currentSql = $execPart
            }
            continue
        }

        # EXECUTE IMMEDIATE 継続行
        if ($state -eq [PlSqlParserState]::InExecuteImmediate) {
            if ($trimmed -match ';\s*$') {
                $currentSql += ' ' + ($trimmed -replace ';\s*$', '')
                $sql = Extract-PlSqlDynamicSql -Expression $currentSql
                if ($sql) {
                    $stmt = [SqlStatement]::new()
                    $stmt.Sql = $sql
                    $stmt.Type = Get-SqlType -SqlText $sql
                    $stmt.Category = 'Dynamic'
                    $stmt.StartLine = $startLine
                    $stmt.EndLine = $lineNum
                    $stmt.SourceFile = $fileName
                    $sqlStatements.Add($stmt)
                }
                $state = [PlSqlParserState]::Normal
                $currentSql = ''
            }
            else {
                $currentSql += ' ' + $trimmed
            }
            continue
        }

        # ================================================
        # DBMS_SQL.PARSE() の検出
        # ================================================
        if ($trimmed -match '(?i)DBMS_SQL\.PARSE\s*\(\s*\w+\s*,\s*(.+?)\s*,') {
            $parseSql = $Matches[1]
            $sql = Extract-PlSqlDynamicSql -Expression $parseSql
            if ($sql) {
                $stmt = [SqlStatement]::new()
                $stmt.Sql = $sql
                $stmt.Type = Get-SqlType -SqlText $sql
                $stmt.Category = 'Dynamic'
                $stmt.StartLine = $lineNum
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $sqlStatements.Add($stmt)
            }
            continue
        }

        # ================================================
        # OPEN cursor FOR の検出
        # ================================================
        if ($trimmed -match '(?i)^OPEN\s+\w+\s+FOR\s+(.+)') {
            $openForPart = $Matches[1]
            $startLine = $lineNum

            # 文字列リテラルの場合（動的SQL）
            if ($openForPart -match "^'") {
                $sql = Extract-PlSqlDynamicSql -Expression $openForPart.TrimEnd(';').Trim()
                $category = 'Dynamic'
            }
            else {
                # 直接SQL文の場合（静的SQL）
                $sql = $openForPart.TrimEnd(';').Trim()
                # 複数行にまたがる可能性
                while (-not $sql.EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                    $i++
                    $sql += ' ' + $lines[$i].Trim().TrimEnd(';')
                }
                $category = 'Static'
            }

            if ($sql) {
                $stmt = [SqlStatement]::new()
                $stmt.Sql = $sql
                $stmt.Type = Get-SqlType -SqlText $sql
                $stmt.Category = $category
                $stmt.StartLine = $startLine
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $sqlStatements.Add($stmt)
            }
            continue
        }

        # ================================================
        # 動的SQL変数への代入の検出
        # v_sql := 'SELECT ...'; / v_sql := v_sql || '...';
        # ================================================
        if ($trimmed -match "(?i)^(\w+)\s*:=\s*(.+?)\s*;\s*$") {
            $varName = $Matches[1]
            $assignExpr = $Matches[2]

            # SQL文字列リテラルを含む代入
            if ($assignExpr -match "'") {
                $sqlPart = Extract-PlSqlStringLiterals -Expression $assignExpr

                if ($sqlPart -and $sqlPart -match '(?i)^\s*(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)') {
                    # 新規代入
                    $dynamicSqlVars[$varName] = @{
                        Fragments = [System.Collections.Generic.List[string]]::new()
                        StartLine = $lineNum
                    }
                    $dynamicSqlVars[$varName].Fragments.Add($sqlPart)
                    $lastFragmentsList = $dynamicSqlVars[$varName].Fragments
                }
                elseif ($dynamicSqlVars.ContainsKey($varName) -or
                        $assignExpr -match "(?i)^$varName\s*\|\|") {
                    # 追記代入
                    if (-not $dynamicSqlVars.ContainsKey($varName)) {
                        $dynamicSqlVars[$varName] = @{
                            Fragments = [System.Collections.Generic.List[string]]::new()
                            StartLine = $lineNum
                        }
                    }
                    $dynamicSqlVars[$varName].Fragments.Add($sqlPart)
                    $lastFragmentsList = $dynamicSqlVars[$varName].Fragments
                }
            }
            continue
        }

        # ================================================
        # CURSOR宣言内のSELECT文
        # ================================================
        if ($trimmed -match '(?i)^CURSOR\s+\w+\s+IS\s+(.+)') {
            $cursorSql = $Matches[1]
            $startLine = $lineNum

            # 複数行にまたがる場合
            while (-not $cursorSql.TrimEnd().EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                $i++
                $lineNum = $i + 1
                $cursorSql += ' ' + $lines[$i].Trim()
            }
            $cursorSql = $cursorSql.TrimEnd(';').Trim()

            $stmt = [SqlStatement]::new()
            $stmt.Sql = $cursorSql
            $stmt.Type = Get-SqlType -SqlText $cursorSql
            $stmt.Category = 'Static'
            $stmt.StartLine = $startLine
            $stmt.EndLine = $lineNum
            $stmt.SourceFile = $fileName
            $sqlStatements.Add($stmt)
            continue
        }

        # ================================================
        # 静的SQL文の検出
        # ================================================
        if ($trimmed -match '(?i)^(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP|TRUNCATE)\b') {
            $startLine = $lineNum
            $staticSql = $trimmed

            # 複数行にまたがるSQL文を収集
            while (-not $staticSql.TrimEnd().EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                $i++
                $lineNum = $i + 1
                $nextLine = $lines[$i].Trim()

                # PL/SQL制御構文に到達したら終了（CASE式のENDは除く）
                # END; / END name; / END LOOP; / END IF; はPL/SQLブロック終端
                if ($nextLine -match '(?i)^(BEGIN|IF|ELSIF|ELSE|LOOP|FOR|WHILE|EXCEPTION|RETURN|DECLARE)\b' -or
                    $nextLine -match '(?i)^END\s*(\w+\s*)?;') {
                    $i--
                    $lineNum = $i + 1
                    break
                }
                $staticSql += ' ' + $nextLine
            }
            $staticSql = $staticSql.TrimEnd(';').Trim()

            $stmt = [SqlStatement]::new()
            $stmt.Sql = $staticSql
            $stmt.Type = Get-SqlType -SqlText $staticSql
            $stmt.Category = 'Static'
            $stmt.StartLine = $startLine
            $stmt.EndLine = $lineNum
            $stmt.SourceFile = $fileName
            $sqlStatements.Add($stmt)
            continue
        }
    }

    # ================================================
    # 動的SQL変数をSQL文として確定
    # ================================================
    foreach ($varEntry in $dynamicSqlVars.GetEnumerator()) {
        $varInfo = $varEntry.Value
        if ($varInfo.Fragments.Count -gt 0) {
            # Fragments には本文 + IF分岐断片が既に含まれている
            $mergedSql = Merge-DynamicSql -Fragments $varInfo.Fragments.ToArray()
            $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'plsql'

            $stmt = [SqlStatement]::new()
            $stmt.Sql = $mergedSql
            $stmt.Type = Get-SqlType -SqlText $mergedSql
            $stmt.Category = 'Dynamic'
            $stmt.StartLine = $varInfo.StartLine
            $stmt.EndLine = $lines.Count
            $stmt.SourceFile = $fileName
            $sqlStatements.Add($stmt)
        }
    }

    $staticCount = ($sqlStatements | Where-Object { $_.Category -eq 'Static' }).Count
    $dynamicCount = ($sqlStatements | Where-Object { $_.Category -eq 'Dynamic' }).Count
    Write-Log -Level INFO -Message "Found $($sqlStatements.Count) SQL statements ($staticCount static, $dynamicCount dynamic)" -LogFile $LogFile

    return $sqlStatements.ToArray()
}

# ============================================================
# Extract-PlSqlStringLiterals: 文字列リテラル部分のみ抽出
# ============================================================
function Extract-PlSqlStringLiterals {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Expression
    )

    $fragments = [System.Collections.Generic.List[string]]::new()
    $expr = $Expression

    # 変数名への参照を削除し、文字列リテラルのみ抽出
    # パターン: 'text' || variable || 'text'
    $parts = $expr -split '\|\|'

    foreach ($part in $parts) {
        $p = $part.Trim()

        # 文字列リテラル
        if ($p -match "^'(.*)'$") {
            $literal = $Matches[1]
            # PL/SQLのエスケープ: '' → '
            $literal = $literal -replace "''", "'"
            $fragments.Add($literal)
        }
        # 変数名 → プレースホルダ
        elseif ($p -match '^[a-zA-Z_][a-zA-Z0-9_.]*$') {
            $fragments.Add("/*:$p*/")
        }
        # 変数 + 追加の文字列を含む部分
        elseif ($p -match "^(\w+)\s*\|\|") {
            $fragments.Add("/*:$($Matches[1])*/")
        }
    }

    if ($fragments.Count -gt 0) {
        return ($fragments -join '')
    }
    return $null
}

# ============================================================
# Extract-PlSqlDynamicSql: 動的SQL式全体の解析
# ============================================================
function Extract-PlSqlDynamicSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Expression
    )

    $expr = $Expression.Trim()

    # USING句を除去
    $expr = $expr -replace '(?i)\s+USING\s+.*$', ''

    # INTO句を除去（EXECUTE IMMEDIATE用）
    $expr = $expr -replace '(?i)\s+INTO\s+\w+.*$', ''

    # 単純な文字列リテラル
    if ($expr -match "^'(.+)'$") {
        $sql = $Matches[1] -replace "''", "'"
        return $sql
    }

    # 変数名（追跡済みの動的SQL変数を参照）
    if ($expr -match '^[a-zA-Z_][a-zA-Z0-9_.]*$') {
        return $null  # 変数参照は動的SQL変数追跡で処理
    }

    # 連結式
    return Extract-PlSqlStringLiterals -Expression $expr
}

# モジュールエクスポート
Export-ModuleMember -Function @(
    'Invoke-PlSqlParser'
)
