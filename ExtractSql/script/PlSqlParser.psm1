#Requires -Version 5.1
<#
.SYNOPSIS
    PL/SQLソースコードからSQL文を抽出するパーサーモジュール
.DESCRIPTION
    静的SQL、動的SQL（EXECUTE IMMEDIATE, DBMS_SQL, OPEN FOR）、
    IF分岐内のSQL断片をすべて抽出・展開します。
#>

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
# Remove-PlSqlInlineComment: 行末インラインコメントを除去
# 文字列リテラル外の -- コメントのみ除去する
# ============================================================
function Remove-PlSqlInlineComment {
    param([string]$Line)

    $result = [System.Text.StringBuilder]::new()
    $inString = $false
    $len = $Line.Length

    for ($k = 0; $k -lt $len; $k++) {
        $ch = $Line[$k]

        if ($inString) {
            # 文字列内: '' はエスケープされたシングルクォート
            if ($ch -eq "'" -and ($k + 1) -lt $len -and $Line[$k + 1] -eq "'") {
                [void]$result.Append("''")
                $k++
            }
            elseif ($ch -eq "'") {
                [void]$result.Append($ch)
                $inString = $false
            }
            else {
                [void]$result.Append($ch)
            }
        }
        else {
            # 文字列外: -- を検出したらそこで終了
            if ($ch -eq '-' -and ($k + 1) -lt $len -and $Line[$k + 1] -eq '-') {
                break
            }
            elseif ($ch -eq "'") {
                [void]$result.Append($ch)
                $inString = $true
            }
            else {
                [void]$result.Append($ch)
            }
        }
    }

    return $result.ToString().TrimEnd()
}

# ============================================================
# Remove-PlSqlBlockComment: ブロックコメントを行単位で除去
# $InBlockComment: 現在のブロックコメント継続状態（[ref] で更新）
# 戻り値: ブロックコメント除去後の文字列（ヒント句 /*+ は保持）
# ============================================================
function Remove-PlSqlBlockComment {
    param(
        [string]$Line,
        [ref]$InBlockComment
    )

    $result = [System.Text.StringBuilder]::new()
    $i = 0
    $len = $Line.Length

    while ($i -lt $len) {
        if ($InBlockComment.Value) {
            # ブロックコメント内: */ を探して終了
            if ($i + 1 -lt $len -and $Line[$i] -eq '*' -and $Line[$i + 1] -eq '/') {
                $InBlockComment.Value = $false
                $i += 2
            } else {
                $i++
            }
            continue
        }

        # /* の検出
        if ($i + 1 -lt $len -and $Line[$i] -eq '/' -and $Line[$i + 1] -eq '*') {
            # ヒント句 /*+ は保持
            if ($i + 2 -lt $len -and $Line[$i + 2] -eq '+') {
                $closeIdx = $Line.IndexOf('*/', $i + 2)
                if ($closeIdx -ge 0) {
                    [void]$result.Append($Line.Substring($i, $closeIdx - $i + 2))
                    $i = $closeIdx + 2
                } else {
                    # ヒント句が次行に続く（稀）: 以降をそのまま保持
                    [void]$result.Append($Line.Substring($i))
                    $i = $len
                }
                continue
            }
            # 通常のブロックコメント: 同一行内で */ を探す
            $closeIdx = $Line.IndexOf('*/', $i + 2)
            if ($closeIdx -ge 0) {
                # 同一行で完結 → スキップして */ の後から続ける
                $i = $closeIdx + 2
            } else {
                # 次行以降に続く → フラグをセットして行の残りを破棄
                $InBlockComment.Value = $true
                $i = $len
            }
            continue
        }

        [void]$result.Append($Line[$i])
        $i++
    }

    return $result.ToString()
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
        [string]$Encoding = 'Default',

        [Parameter()]
        [string]$LogFile = ''
    )

    $fileName = [System.IO.Path]::GetFileName($FilePath)
    Write-Log -Level INFO -Message "Processing: $fileName (PL/SQL)" -LogFile $LogFile

    $detectedEncoding = Get-FileEncoding -FilePath $FilePath -FallbackEncoding $Encoding
    Write-Log -Level INFO -Message "  Encoding: $detectedEncoding (fallback: $Encoding)" -LogFile $LogFile
    $lines = Get-Content -Path $FilePath -Encoding $detectedEncoding
    $sqlStatements = [System.Collections.Generic.List[object]]::new()

    $state = [PlSqlParserState]::Normal
    $currentSql = ''
    $startLine = 0
    $inBlockComment = $false
    $dynamicSqlVars = @{}  # 変数名 → @{Fragments; StartLine} のマッピング

    # 最後に更新された動的SQL変数の Fragments リストへの参照。
    # IF分岐の断片をその変数に直接追加するために使用する。
    $lastFragmentsList = $null
    $lastVarName = $null

    # EXCEPTIONブロック追跡: ハンドラ内の変数代入を動的SQLと誤認しないためのフラグ
    $inExceptionBlock = $false
    $beginNestLevel = 0
    $exceptionNestLevel = -1

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lineNum = $i + 1
        $line = $lines[$i]
        $trimmed = $line.Trim()

        # ================================================
        # ブロックコメントのスキップ（/*+ ヒント句は除外）
        # Remove-PlSqlBlockComment で1行完結・複数行・行中開始を統一処理する
        # ================================================
        $trimmed = (Remove-PlSqlBlockComment -Line $trimmed -InBlockComment ([ref]$inBlockComment)).Trim()
        if ($trimmed -eq '') { continue }

        # 行コメントのスキップ
        if ($trimmed -match '^--') {
            continue
        }

        # ================================================
        # BEGIN/ENDネストレベルの追跡とEXCEPTIONブロック検出
        # ================================================
        if ($trimmed -match '(?i)^BEGIN\b') {
            $beginNestLevel++
        }
        if ($trimmed -match '(?i)^END\b') {
            if ($inExceptionBlock -and $beginNestLevel -le $exceptionNestLevel) {
                $inExceptionBlock = $false
                $exceptionNestLevel = -1
            }
            $beginNestLevel--
            if ($beginNestLevel -lt 0) { $beginNestLevel = 0 }
        }
        if ($trimmed -match '(?i)^EXCEPTION\b') {
            $inExceptionBlock = $true
            $exceptionNestLevel = $beginNestLevel
            continue
        }

        # ================================================
        # IF分岐の検出と展開
        # ================================================
        if ($trimmed -match '(?i)^\s*IF\s+(.+?)\s+THEN') {
            # IF ブロック全体を収集（空白行・空行は除外してから渡す）
            $ifLines = [System.Collections.Generic.List[string]]::new()
            $ifStartLine = $lineNum
            $ifNestLevel = 0

            for ($j = $i; $j -lt $lines.Count; $j++) {
                $ifLine = $lines[$j].Trim()
                # 空行はスキップ（Expand-IfBranchesのMandatory[string[]]パラメータが空文字列を拒否するため）
                if ($ifLine -ne '') {
                    $ifLines.Add($lines[$j])
                }

                if ($ifLine -match '(?i)^\s*IF\s+') {
                    $ifNestLevel++
                }
                if ($ifLine -match '(?i)^\s*END\s+IF') {
                    $ifNestLevel--
                    if ($ifNestLevel -eq 0) {
                        break
                    }
                }
                # PROCEDURE/FUNCTION/PACKAGE 境界に達したらスキャンを停止
                # （END IF なしの不正なIFブロックが次のオブジェクト定義を飲み込まないようにする）
                if ($j -gt $i -and $ifLine -match '(?i)^\s*(PROCEDURE|FUNCTION|PACKAGE)\b') {
                    $j--  # この行を外側ループで再処理させる
                    break
                }
            }
            # END IFが見つからずループが完了した場合、$jは$lines.Countになる（配列範囲外）
            # インナースキャンで$lines[$j]がnullになりエラーになるため、クランプする
            if ($j -ge $lines.Count) {
                $j = $lines.Count - 1
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

            if ($ifLines.Count -gt 0) {
                $branchResults = Expand-IfBranches -Lines $ifLines.ToArray() `
                    -Language 'plsql' -ExtractSqlFromLine $extractSqlFromLine

                # ブランチコメント（"-- [Branch N] ..."）を除いた実際のSQL断片のみチェックする
                # Expand-IfBranchesは分岐があれば必ずコメント行を返すため、
                # コメント行だけの場合は動的SQL断片なし（静的SQLのIF分岐）と判断する
                $realFragments = @($branchResults | Where-Object { $_ -notmatch '^-- \[Branch \d+\]' })
                if ($realFragments.Count -gt 0) {
                    # 直前に操作した動的SQL変数のFragmentsリストに直接追加する。
                    # 変数が未確定の場合は後続の変数確定フェーズで拾われるようにフォールバック。
                    if ($lastFragmentsList) {
                        foreach ($fragment in $branchResults) {
                            $lastFragmentsList.Add($fragment)
                        }
                        if ($lastVarName -and $dynamicSqlVars.ContainsKey($lastVarName)) {
                            $dynamicSqlVars[$lastVarName].EndLine = $j + 1
                        }
                    }
                    # 変数が特定できない場合は破棄（ログ警告を出す）
                    else {
                        Write-Log -Level WARN -Message "Line ${ifStartLine}: IF分岐の断片を関連付ける動的SQL変数が見つかりません" -LogFile $LogFile
                    }
                }
            }

            # IFブロック内の直接静的SQL文（UPDATE/DELETE/INSERT/SELECT等）を抽出する
            $m = $i
            while ($m -le $j) {
                $mTrimmed = $lines[$m].Trim()
                $mLineNum = $m + 1

                # 空行・行コメント・制御構文はスキップ
                if ($mTrimmed -eq '' -or
                    $mTrimmed -match '^--' -or
                    $mTrimmed -match '(?i)^(IF|ELSIF|ELSE|END\s+IF|BEGIN|END)\b') {
                    $m++
                    continue
                }

                if ($mTrimmed -match '(?i)^(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP|TRUNCATE)\b') {
                    $innerSqlStart = $mLineNum
                    $innerSql = Remove-PlSqlInlineComment -Line $mTrimmed
                    $m++

                    # CASE式のネストレベルを追跡（CASE内のELSEをPL/SQL ELSE と誤検知しないため）
                    $innerCaseNestLevel = ([regex]::Matches($innerSql, '(?i)\bCASE\b')).Count - `
                                         ([regex]::Matches($innerSql, '(?i)\bEND\b')).Count
                    if ($innerCaseNestLevel -lt 0) { $innerCaseNestLevel = 0 }

                    while ($m -le $j -and -not $innerSql.TrimEnd().EndsWith(';')) {
                        $rawMLine = $lines[$m].Trim()
                        $rawMLine = Remove-PlSqlBlockComment -Line $rawMLine -InBlockComment ([ref]$inBlockComment)
                        if ($rawMLine.Trim() -eq '') { $m++; continue }
                        $nextMLine = Remove-PlSqlInlineComment -Line $rawMLine
                        if ($nextMLine -match '(?i)^(IF|ELSIF|LOOP|WHILE|EXCEPTION|RETURN|DECLARE)\b' -or
                            ($nextMLine -match '(?i)^FOR\b' -and $nextMLine -notmatch '(?i)^FOR\s+UPDATE\b') -or
                            ($nextMLine -match '(?i)^ELSE\s*$' -and $innerCaseNestLevel -le 0) -or
                            ($nextMLine -match '(?i)^END\b' -and $innerCaseNestLevel -le 0) -or
                            $nextMLine -match '^--') {
                            break
                        }

                        # CASE/ENDネストレベルを更新
                        $innerCaseNestLevel += ([regex]::Matches($nextMLine, '(?i)\bCASE\b')).Count
                        $innerCaseNestLevel -= ([regex]::Matches($nextMLine, '(?i)\bEND\b')).Count
                        if ($innerCaseNestLevel -lt 0) { $innerCaseNestLevel = 0 }

                        $innerSql += "`n" + $nextMLine
                        $m++
                    }
                    $innerSql = $innerSql.TrimEnd(';').Trim()

                    if ($innerSql) {
                        $stmt = New-SqlStatement
                        $stmt.Sql = $innerSql
                        $stmt.Type = Get-SqlType -SqlText $innerSql
                        $stmt.Category = 'Static'
                        $stmt.StartLine = $innerSqlStart
                        $stmt.EndLine = $m
                        $stmt.SourceFile = $fileName
                        $sqlStatements.Add($stmt)
                    }
                }
                else {
                    $m++
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
                    $stmt = New-SqlStatement
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
                    $stmt = New-SqlStatement
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
                $stmt = New-SqlStatement
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
                # USING句を除去してFOR直後の内容を確認する
                $openForContent = ($openForPart -replace '(?i)\s+USING\s+.*$', '' -replace ';\s*$', '').Trim()

                # 変数参照のみの場合（OPEN c FOR v_sql / OPEN c FOR v_sql USING ...）は
                # 動的SQL変数追跡で処理済みのためスキップする
                if ($openForContent -notmatch '(?i)^(SELECT|INSERT|UPDATE|DELETE|MERGE|WITH)\b') {
                    continue
                }

                # 直接SQL文の場合（静的SQL）
                $sql = $openForPart.TrimEnd(';').Trim()
                # 複数行にまたがる可能性
                while (-not $sql.EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                    $i++
                    $rawOpenLine = $lines[$i].Trim()
                    # ブロックコメント除去（/*+ ヒント句は保持）
                    $rawOpenLine = Remove-PlSqlBlockComment -Line $rawOpenLine -InBlockComment ([ref]$inBlockComment)
                    if ($rawOpenLine.Trim() -eq '') { continue }
                    $nextOpenLine = Remove-PlSqlInlineComment -Line $rawOpenLine
                    $sql += "`n" + $nextOpenLine.TrimEnd(';')
                }
                $category = 'Static'
            }

            if ($sql) {
                $stmt = New-SqlStatement
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
        if (-not $inExceptionBlock -and $trimmed -match "(?i)^(\w+)\s*:=\s*(.+?)\s*;\s*$") {
            $varName = $Matches[1]
            $assignExpr = $Matches[2]

            # SQL文字列リテラルを含む代入
            if ($assignExpr -match "'") {
                $sqlPart = Extract-PlSqlStringLiterals -Expression $assignExpr

                if ($sqlPart -and $sqlPart -match '(?i)^\s*(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                    # 同名変数に既存の断片がある場合は先に確定させる（複数プロシージャで同名変数が使われる場合の対応）
                    if ($dynamicSqlVars.ContainsKey($varName) -and $dynamicSqlVars[$varName].Fragments.Count -gt 0) {
                        $prevInfo = $dynamicSqlVars[$varName]
                        $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                        $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'plsql'
                        $prevStmt = New-SqlStatement
                        $prevStmt.Sql = $prevMerged
                        $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                        $prevStmt.Category = 'Dynamic'
                        $prevStmt.StartLine = $prevInfo.StartLine
                        $prevStmt.EndLine = $lineNum - 1
                        $prevStmt.SourceFile = $fileName
                        $sqlStatements.Add($prevStmt)
                    }
                    # 新規代入
                    $dynamicSqlVars[$varName] = @{
                        Fragments = [System.Collections.Generic.List[string]]::new()
                        StartLine = $lineNum
                        EndLine   = $lineNum
                    }
                    $dynamicSqlVars[$varName].Fragments.Add($sqlPart)
                    $lastFragmentsList = $dynamicSqlVars[$varName].Fragments
                    $lastVarName = $varName
                }
                elseif ($dynamicSqlVars.ContainsKey($varName) -or
                        ($assignExpr -match "(?i)^$varName\s*\|\|" -and
                         $sqlPart -and
                         ($sqlPart -replace '/\*:[^*]+\*/', '').Trim() -match '(?i)^\s*(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b')) {
                    # 追記代入
                    if (-not $dynamicSqlVars.ContainsKey($varName)) {
                        $dynamicSqlVars[$varName] = @{
                            Fragments = [System.Collections.Generic.List[string]]::new()
                            StartLine = $lineNum
                            EndLine   = $lineNum
                        }
                    }
                    $dynamicSqlVars[$varName].Fragments.Add($sqlPart)
                    $dynamicSqlVars[$varName].EndLine = $lineNum
                    $lastFragmentsList = $dynamicSqlVars[$varName].Fragments
                    $lastVarName = $varName
                }
            }
            continue
        }

        # ================================================
        # CURSOR宣言内のSELECT文
        # CURSOR name [(params)] IS [sql]
        # 引数あり/なし・IS が同一行/次行のケースに対応
        # ================================================
        if ($trimmed -match '(?i)^CURSOR\s+(\w+)') {
            $startLine = $lineNum
            $cursorName = $Matches[1]
            $cursorSql = ''
            $foundIs = $false

            # 同一行に IS がある場合: CURSOR name IS ... / CURSOR name(params) IS ...
            if ($trimmed -match '(?i)\bIS\s+(.+)$') {
                $cursorSql = Remove-PlSqlInlineComment -Line $Matches[1]
                $foundIs = $true
            }
            elseif ($trimmed -match '(?i)\bIS\s*$') {
                $foundIs = $true
                if (($i + 1) -lt $lines.Count) {
                    $i++
                    $lineNum = $i + 1
                    $cursorSql = Remove-PlSqlInlineComment -Line $lines[$i].Trim()
                }
                else { continue }
            }
            else {
                # IS が次行以降にある場合（引数リストが複数行にわたるケース）
                $lookIdx = $i + 1
                while ($lookIdx -lt $lines.Count) {
                    $lookLine = $lines[$lookIdx].Trim()
                    if ($lookLine -match '(?i)^(BEGIN|FUNCTION|PROCEDURE|END|DECLARE)\b') { break }
                    if ($lookLine -match '(?i)^IS\s+(.+)$') {
                        $cursorSql = Remove-PlSqlInlineComment -Line $Matches[1]
                        $i = $lookIdx; $lineNum = $i + 1; $foundIs = $true; break
                    }
                    elseif ($lookLine -match '(?i)^IS\s*$') {
                        $i = $lookIdx; $lineNum = $i + 1; $foundIs = $true
                        if (($i + 1) -lt $lines.Count) {
                            $i++; $lineNum = $i + 1
                            $cursorSql = Remove-PlSqlInlineComment -Line $lines[$i].Trim()
                        }
                        break
                    }
                    $lookIdx++
                }
            }

            if (-not $foundIs) { continue }

            # 複数行にまたがる場合
            while (-not $cursorSql.TrimEnd().EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                $i++
                $lineNum = $i + 1
                $rawCursorLine = $lines[$i].Trim()

                # ブロックコメント除去（/*+ ヒント句は保持、1行完結・複数行・行中開始すべて対応）
                $rawCursorLine = Remove-PlSqlBlockComment -Line $rawCursorLine -InBlockComment ([ref]$inBlockComment)
                if ($rawCursorLine.Trim() -eq '') { continue }

                $nextCursorLine = Remove-PlSqlInlineComment -Line $rawCursorLine

                # PL/SQL制御構文に到達したら終了
                # FOR UPDATE / FOR UPDATE OF は SQL のロック句なので FOR ループと区別する
                if ($nextCursorLine -match '(?i)^(BEGIN|IF|ELSIF|LOOP|WHILE|EXCEPTION|RETURN|DECLARE)\b' -or
                    ($nextCursorLine -match '(?i)^FOR\b' -and $nextCursorLine -notmatch '(?i)^FOR\s+UPDATE\b') -or
                    $nextCursorLine -match '(?i)^ELSE\s*$' -or
                    $nextCursorLine -match '(?i)^END\s*(\w+\s*)?;') {
                    $i--
                    $lineNum = $i + 1
                    break
                }
                $cursorSql += "`n" + $nextCursorLine
            }
            $cursorSql = $cursorSql.TrimEnd(';').Trim()

            if ($cursorSql) {
                $stmt = New-SqlStatement
                $stmt.Sql = $cursorSql
                $stmt.Type = Get-SqlType -SqlText $cursorSql
                $stmt.Category = 'Static'
                $stmt.StartLine = $startLine
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $stmt.CursorName = $cursorName
                $sqlStatements.Add($stmt)
            }
            continue
        }

        # ================================================
        # 静的SQL文の検出
        # ================================================
        # PL/SQLオブジェクト宣言（PACKAGE/PROCEDURE/FUNCTION/TRIGGER/TYPE/BODY）はスキップ
        if ($trimmed -match '(?i)^CREATE\s+(OR\s+REPLACE\s+)?(PACKAGE|PROCEDURE|FUNCTION|TRIGGER|TYPE)\b') {
            continue
        }
        if ($trimmed -match '(?i)^(SELECT|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP|TRUNCATE)\b') {
            $startLine = $lineNum
            # 先頭行もインラインコメントを除去してから使用
            $staticSql = Remove-PlSqlInlineComment -Line $trimmed

            # CASE式のネストレベルを追跡（CASE内のELSEをPL/SQL ELSE と誤検知しないため）
            $caseNestLevel = ([regex]::Matches($staticSql, '(?i)\bCASE\b')).Count - `
                             ([regex]::Matches($staticSql, '(?i)\bEND\b')).Count
            if ($caseNestLevel -lt 0) { $caseNestLevel = 0 }

            # 複数行にまたがるSQL文を収集
            while (-not $staticSql.TrimEnd().EndsWith(';') -and ($i + 1) -lt $lines.Count) {
                $i++
                $lineNum = $i + 1
                $rawNextLine = $lines[$i].Trim()
                $rawNextLine = Remove-PlSqlBlockComment -Line $rawNextLine -InBlockComment ([ref]$inBlockComment)
                if ($rawNextLine.Trim() -eq '') { continue }
                $nextLine = Remove-PlSqlInlineComment -Line $rawNextLine

                # PL/SQL制御構文に到達したら終了
                # CASE式内の ELSE はネストレベルが1以上の場合はスキップ
                # FOR UPDATE / FOR UPDATE OF は SQL のロック句なので FOR ループと区別する
                if ($nextLine -match '(?i)^(BEGIN|IF|ELSIF|LOOP|WHILE|EXCEPTION|RETURN|DECLARE)\b' -or
                    ($nextLine -match '(?i)^FOR\b' -and $nextLine -notmatch '(?i)^FOR\s+UPDATE\b') -or
                    ($nextLine -match '(?i)^ELSE\s*$' -and $caseNestLevel -le 0) -or
                    $nextLine -match '(?i)^END\s*(\w+\s*)?;') {
                    $i--
                    $lineNum = $i + 1
                    break
                }

                # CASE/ENDネストレベルを更新
                $caseNestLevel += ([regex]::Matches($nextLine, '(?i)\bCASE\b')).Count
                $caseNestLevel -= ([regex]::Matches($nextLine, '(?i)\bEND\b')).Count
                if ($caseNestLevel -lt 0) { $caseNestLevel = 0 }

                $staticSql += "`n" + $nextLine
            }
            $staticSql = $staticSql.TrimEnd(';').Trim()

            $stmt = New-SqlStatement
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

            $stmt = New-SqlStatement
            $stmt.Sql = $mergedSql
            $stmt.Type = Get-SqlType -SqlText $mergedSql
            $stmt.Category = 'Dynamic'
            $stmt.StartLine = $varInfo.StartLine
            $stmt.EndLine = $varInfo.EndLine
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
    $i = 0
    $len = $Expression.Length

    while ($i -lt $len) {
        $ch = $Expression[$i]

        if ($ch -eq "'") {
            # 文字列リテラル: 閉じクォートまで読む（'' はエスケープ）
            $i++
            $sb = [System.Text.StringBuilder]::new()
            while ($i -lt $len) {
                if ($Expression[$i] -eq "'" -and ($i + 1) -lt $len -and $Expression[$i + 1] -eq "'") {
                    [void]$sb.Append("'")
                    $i += 2
                } elseif ($Expression[$i] -eq "'") {
                    $i++
                    break
                } else {
                    [void]$sb.Append($Expression[$i])
                    $i++
                }
            }
            $fragments.Add($sb.ToString())
        } elseif ([char]::IsLetter($ch) -or $ch -eq '_') {
            # 識別子（変数名）: 英数字・アンダースコア・ドットを読む
            $start = $i
            while ($i -lt $len -and ([char]::IsLetterOrDigit($Expression[$i]) -or $Expression[$i] -eq '_' -or $Expression[$i] -eq '.')) {
                $i++
            }
            $name = $Expression.Substring($start, $i - $start)
            $fragments.Add("/*:$name*/")
        } else {
            # 演算子（||）・空白などをスキップ
            $i++
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

    # 文字列リテラルを最初に判定する（USING/INTO除去より前）
    # これにより 'INSERT INTO ...' の INTO がSQL内部のものと誤って除去されるのを防ぐ
    # パターン: 'SQL'  または  'SQL' INTO var  または  'SQL' USING ...
    if ($expr -match "^'((?:[^']|'')*)'(?:\s+(?i:INTO|USING)\s+.*)?$") {
        $sql = $Matches[1] -replace "''", "'"
        return $sql
    }

    # 変数式・連結式に対してのみ USING/INTO 句を除去
    $expr = $expr -replace '(?i)\s+USING\s+.*$', ''
    $expr = $expr -replace '(?i)\s+INTO\s+\w+.*$', ''

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
