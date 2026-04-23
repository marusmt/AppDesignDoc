#Requires -Version 5.1
<#
.SYNOPSIS
    VB.NETソースコードからSQL文を抽出するパーサーモジュール
.DESCRIPTION
    CommandText代入、StringBuilder、文字列連結（& / +）、
    String.Format、補間文字列、If分岐展開に対応します。
    行継続文字（_）による複数行連結にも対応しています。
#>

# ============================================================
# 状態管理用Enum
# ============================================================
enum VbNetParserState {
    Normal
    InSqlAssign
    InStringBuilder
    InCommandText
    InIfBlock
}

# ============================================================
# Invoke-VbNetParser: VB.NETパース実行
# ============================================================
function Invoke-VbNetParser {
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
    Write-Log -Level INFO -Message "Processing: $fileName (VB.NET)" -LogFile $LogFile

    $detectedEncoding = Get-FileEncoding -FilePath $FilePath -FallbackEncoding $Encoding
    Write-Log -Level INFO -Message "  Encoding: $detectedEncoding (fallback: $Encoding)" -LogFile $LogFile
    $rawLines = Get-Content -Path $FilePath -Encoding $detectedEncoding

    # 行継続文字（_）を事前に結合し、元の行番号マッピングも取得する
    $joined = Join-VbNetContinuationLines -Lines $rawLines
    $lines = $joined.Lines
    $originalLineNumbers = $joined.OriginalLineNumbers

    $sqlStatements = [System.Collections.Generic.List[object]]::new()
    $dynamicSqlVars = @{}   # 変数名 → @{Fragments; StartLine; EndLine}
    $sbVars = @{}           # StringBuilder変数名 → @{Fragments; StartLine; EndLine}

    # 最後に更新された動的SQL変数/SBの Fragments リストへの参照。
    # IF分岐の断片をその変数に直接追加するために使用する。
    $lastFragmentsList = $null
    $lastVarName = $null      # 最後に更新された変数名
    $lastVarSource = $null    # 最後に更新されたハッシュテーブル（$dynamicSqlVars または $sbVars）
    $currentWithVar = $null   # With ブロックで対象となっている変数名

    $currentMethodName = ''   # 現在処理中のメソッド名（Sub/Function）

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $lineNum = $originalLineNumbers[$i]
        $line = $lines[$i]
        $trimmed = $line.Trim()

        # ================================================
        # コメントのスキップ・空行スキップ
        # ================================================
        if ($trimmed -eq '') {
            continue
        }
        if ($trimmed -match "^\s*'") {
            continue
        }
        # Fix: -match は2引数を受け取れないため (?i) インラインフラグを使用
        if ($trimmed -match '(?i)^\s*REM\s') {
            continue
        }

        # インラインコメント除去
        $trimmed = Remove-VbNetInlineComment -Line $trimmed

        # ================================================
        # With ブロックの追跡
        # ================================================
        # With varName → With ブロック開始
        if ($trimmed -match '(?i)^With\s+(\w+)\s*$') {
            $currentWithVar = $Matches[1]
            continue
        }
        # End With → With ブロック終了
        if ($trimmed -match '(?i)^End\s+With\s*$') {
            $currentWithVar = $null
            continue
        }
        # With ブロック内の ".Append(...)" を "varName.Append(...)" に正規化
        if ($currentWithVar -and $trimmed -match '(?i)^\.\w') {
            $trimmed = $currentWithVar + $trimmed
        }

        # ================================================
        # メソッド（Sub/Function）宣言の検出
        # 例: Public Sub LoadData() / Private Function BuildSql(...) As String
        # End Sub/Function が検出されなかった場合の保険として、蓄積済みSQLをここで確定する
        # ================================================
        if ($trimmed -match '(?i)\b(?:Sub|Function)\s+(\w+)') {
            if ($dynamicSqlVars.Count -gt 0 -or $sbVars.Count -gt 0) {
                foreach ($varEntry in $dynamicSqlVars.GetEnumerator()) {
                    $varInfo = $varEntry.Value
                    if ($varInfo.Fragments.Count -gt 0) {
                        $mergedSql = Merge-DynamicSql -Fragments $varInfo.Fragments.ToArray()
                        $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'
                        if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                            $stmt = New-SqlStatement
                            $stmt.Sql = $mergedSql
                            $stmt.Type = Get-SqlType -SqlText $mergedSql
                            $stmt.Category = 'Dynamic'
                            $stmt.StartLine = $varInfo.StartLine
                            $stmt.EndLine = $varInfo.EndLine
                            $stmt.SourceFile = $fileName
                            $stmt.MethodName = $currentMethodName
                            $sqlStatements.Add($stmt)
                        }
                    }
                }
                foreach ($sbEntry in $sbVars.GetEnumerator()) {
                    $sbInfo = $sbEntry.Value
                    if ($sbInfo.Fragments.Count -gt 0) {
                        $mergedSql = Merge-DynamicSql -Fragments $sbInfo.Fragments.ToArray()
                        $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'
                        if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                            $stmt = New-SqlStatement
                            $stmt.Sql = $mergedSql
                            $stmt.Type = Get-SqlType -SqlText $mergedSql
                            $stmt.Category = 'Dynamic'
                            $stmt.StartLine = $sbInfo.StartLine
                            $stmt.EndLine = $sbInfo.EndLine
                            $stmt.SourceFile = $fileName
                            $stmt.MethodName = $currentMethodName
                            $sqlStatements.Add($stmt)
                        }
                    }
                }
                $dynamicSqlVars = @{}
                $sbVars = @{}
            }
            # 蓄積変数の有無にかかわらず常にリセット（前メソッドの残存参照を確実にクリア）
            $lastFragmentsList = $null
            $lastVarName = $null
            $lastVarSource = $null
            $currentWithVar = $null
            $currentMethodName = $Matches[1]
            continue
        }

        # ================================================
        # メソッド境界の検出: End Sub / End Function
        # 同名ローカル変数が別メソッドで再利用される場合に備え、
        # メソッド終了時に蓄積した SQL 変数を確定してスコープをリセットする。
        # ================================================
        if ($trimmed -match '(?i)^End\s+(Sub|Function)\s*$') {
            foreach ($varEntry in $dynamicSqlVars.GetEnumerator()) {
                $varInfo = $varEntry.Value
                if ($varInfo.Fragments.Count -gt 0) {
                    $mergedSql = Merge-DynamicSql -Fragments $varInfo.Fragments.ToArray()
                    $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'
                    if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                        $stmt = New-SqlStatement
                        $stmt.Sql = $mergedSql
                        $stmt.Type = Get-SqlType -SqlText $mergedSql
                        $stmt.Category = 'Dynamic'
                        $stmt.StartLine = $varInfo.StartLine
                        $stmt.EndLine = $varInfo.EndLine
                        $stmt.SourceFile = $fileName
                        $stmt.MethodName = $currentMethodName
                        $sqlStatements.Add($stmt)
                    }
                }
            }
            foreach ($sbEntry in $sbVars.GetEnumerator()) {
                $sbInfo = $sbEntry.Value
                if ($sbInfo.Fragments.Count -gt 0) {
                    $mergedSql = Merge-DynamicSql -Fragments $sbInfo.Fragments.ToArray()
                    $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'
                    # プレースホルダ /*:...*/ が先頭に来る場合も考慮してSQLキーワードを検出
                    if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                        $stmt = New-SqlStatement
                        $stmt.Sql = $mergedSql
                        $stmt.Type = Get-SqlType -SqlText $mergedSql
                        $stmt.Category = 'Dynamic'
                        $stmt.StartLine = $sbInfo.StartLine
                        $stmt.EndLine = $sbInfo.EndLine
                        $stmt.SourceFile = $fileName
                        $stmt.MethodName = $currentMethodName
                        $sqlStatements.Add($stmt)
                    }
                }
            }
            $dynamicSqlVars = @{}
            $sbVars = @{}
            $lastFragmentsList = $null
            $lastVarName = $null
            $lastVarSource = $null
            $currentWithVar = $null
            $currentMethodName = ''
            continue
        }

        # ================================================
        # If分岐の検出と展開
        # ================================================
        if ($trimmed -match '(?i)^\s*If\s+(.+?)\s+Then\s*$') {
            $ifLines = [System.Collections.Generic.List[string]]::new()
            $ifNestLevel = 0

            for ($j = $i; $j -lt $lines.Count; $j++) {
                $ifLine = $lines[$j].Trim()

                # メソッド境界（End Sub/Function）でスキャンを停止
                # 越えると$jが配列外になりEndLine=0になるのを防ぐ
                if ($ifLine -match '(?i)^End\s+(Sub|Function)\b') {
                    if ($j -gt $i) { $j-- }
                    break
                }

                # 空行はスキップ（Remove-VbNetInlineCommentは空文字列を受け付けないため、
                # 空行チェックの後でインラインコメント除去を行う）
                if ($ifLine -ne '') {
                    $ifLines.Add($lines[$j])
                    # インラインコメント除去後の行をネスト追跡に使用（例: If cond Then 'comment）
                    $ifLineNoComment = Remove-VbNetInlineComment -Line $ifLine

                    if ($ifLineNoComment -match '(?i)^\s*If\s+.+\s+Then\s*$') {
                        $ifNestLevel++
                    }
                    if ($ifLineNoComment -match '(?i)^\s*End\s+If') {
                        $ifNestLevel--
                        if ($ifNestLevel -eq 0) {
                            break
                        }
                    }
                }
            }

            # SQL断片抽出用スクリプトブロック
            $extractSqlFromLine = {
                param($ln)
                $t = $ln.Trim()
                $result = $null

                # StringBuilder.Append にメソッド呼び出しが渡されている場合（部分抽出）
                # 例: sb.Append(BuildWithCteBlock("SELECT A, B, C FROM M_REF_TABLE "))
                if ($t -match '(?i)\.Append(?:Line)?\s*\(\s*([a-zA-Z_][\w.]*\s*\(.*\))\s*\)') {
                    $callExpr = $Matches[1].Trim()
                    return "/*:${callExpr}*/"
                }

                # StringBuilder.Append / .AppendLine パターン
                if ($t -match '(?i)\.Append(?:Line)?\s*\(\s*"(.+?)"\s*\)') {
                    $result = $Matches[1] -replace '""', '"'
                    return $result
                }

                # 文字列連結代入: sql &= "..." / sql += "..."
                if ($t -match '(?i)^\w+\s*[&+]=\s*"(.+?)"') {
                    $result = $Matches[1] -replace '""', '"'
                    return $result
                }

                # 変数代入: sql = sql & "..."
                if ($t -match '(?i)^\w+\s*=\s*\w+\s*&\s*"(.+?)"') {
                    $result = $Matches[1] -replace '""', '"'
                    return $result
                }

                # CommandText代入
                if ($t -match '(?i)\.CommandText\s*=\s*"(.+?)"') {
                    $result = $Matches[1] -replace '""', '"'
                    return $result
                }

                return $null
            }

            if ($ifLines.Count -eq 0) {
                $branchResults = @()
            }
            else {
                # インラインコメントを除去してExpand-IfBranchesに渡す
                # （例: "If cond Then 'comment" が $ifPattern にマッチするよう正規化）
                $ifLinesForExpand = $ifLines.ToArray() | ForEach-Object {
                    Remove-VbNetInlineComment -Line $_
                }
                $branchResults = Expand-IfBranches -Lines $ifLinesForExpand `
                    -Language 'vbnet' -ExtractSqlFromLine $extractSqlFromLine
            }

            if ($branchResults.Count -gt 0) {
                $endLineNum = if ($j -lt $originalLineNumbers.Count) { $originalLineNumbers[$j] } else { $originalLineNumbers[$originalLineNumbers.Count - 1] }

                # $$SQL_RESET:varName$$ センチネルの有無を確認
                # センチネルは Expand-IfBranches が sb = New StringBuilder を検出した境界を示す
                $hasResetSentinel = $branchResults | Where-Object { $_ -match '^\$\$SQL_RESET:' }

                if ($hasResetSentinel) {
                    # ------------------------------------------------
                    # センチネルあり: IF 内で sb = New StringBuilder が複数回実行されるケース
                    # $branchResults をセンチネル境界で分割し、各セグメントを個別 SQL として処理する
                    # ------------------------------------------------
                    $segments = [System.Collections.Generic.List[hashtable]]::new()
                    $segFragments = [System.Collections.Generic.List[string]]::new()
                    $segVarName   = $null  # null = センチネル前（RESETなし）

                    foreach ($br in $branchResults) {
                        if ($br -match '^\$\$SQL_RESET:(\w+)\$\$$') {
                            # 前セグメントを保存
                            $segments.Add(@{ VarName = $segVarName; Fragments = $segFragments.ToArray() })
                            $segVarName   = $Matches[1]
                            $segFragments = [System.Collections.Generic.List[string]]::new()
                        } else {
                            $segFragments.Add($br)
                        }
                    }
                    $segments.Add(@{ VarName = $segVarName; Fragments = $segFragments.ToArray() })

                    for ($si = 0; $si -lt $segments.Count; $si++) {
                        $seg     = $segments[$si]
                        $segVar  = $seg.VarName
                        $segFrag = $seg.Fragments
                        $isLast  = ($si -eq $segments.Count - 1)

                        if ($null -eq $segVar) {
                            # センチネル前: 既存 $lastFragmentsList があれば追加（通常は分岐コメントのみ）
                            if ($lastFragmentsList -and $segFrag.Count -gt 0) {
                                foreach ($fr in $segFrag) { $lastFragmentsList.Add($fr) }
                                if ($lastVarName -and $lastVarSource -and $lastVarSource.ContainsKey($lastVarName)) {
                                    $lastVarSource[$lastVarName].EndLine = $endLineNum
                                }
                            }
                            continue
                        }

                        # RESET 後セグメント: 前の同名変数があれば即時確定して出力する
                        if ($sbVars.ContainsKey($segVar) -and $sbVars[$segVar].Fragments.Count -gt 0) {
                            $prevInfo   = $sbVars[$segVar]
                            $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                            $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                            if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                                $prevStmt            = New-SqlStatement
                                $prevStmt.Sql        = $prevMerged
                                $prevStmt.Type       = Get-SqlType -SqlText $prevMerged
                                $prevStmt.Category   = 'Dynamic'
                                $prevStmt.StartLine  = $prevInfo.StartLine
                                $prevStmt.EndLine    = $endLineNum
                                $prevStmt.SourceFile = $fileName
                                $prevStmt.MethodName = $prevInfo.MethodName
                                $sqlStatements.Add($prevStmt)
                            }
                            $sbVars.Remove($segVar)
                        }

                        # 断片がある場合のみ新しい $sbVars エントリを作成
                        if ($segFrag.Count -gt 0) {
                            $sbVars[$segVar] = @{
                                Fragments  = [System.Collections.Generic.List[string]]::new()
                                StartLine  = $lineNum
                                EndLine    = $endLineNum
                                MethodName = $currentMethodName
                            }
                            foreach ($fr in $segFrag) { $sbVars[$segVar].Fragments.Add($fr) }
                        }

                        # 最後のセグメントを $lastFragmentsList に設定
                        if ($isLast -and $sbVars.ContainsKey($segVar)) {
                            $lastFragmentsList = $sbVars[$segVar].Fragments
                            $lastVarName       = $segVar
                            $lastVarSource     = $sbVars
                        }
                    }
                }
                else {
                    # ------------------------------------------------
                    # センチネルなし: 従来の処理
                    # ------------------------------------------------
                    if ($lastFragmentsList) {
                        foreach ($fragment in $branchResults) {
                            $lastFragmentsList.Add($fragment)
                        }
                        if ($lastVarName -and $lastVarSource -and $lastVarSource.ContainsKey($lastVarName)) {
                            $lastVarSource[$lastVarName].EndLine = $endLineNum
                        }
                    }
                    else {
                        # $lastFragmentsList が null のケース:
                        # IF 分岐のみで構成されるヘルパーメソッドなど、直前に .Append がない場合。
                        # $ifLinesForExpand から変数名を推定して新しい $sbVars エントリを作成する。
                        # 優先度順に変数名を推定:
                        # 1. With ブロック内の If の場合、外側の $currentWithVar を使う
                        #    （IF ボディが ".Append(...)" 先頭ドット形式のとき $ifLinesForExpand からは見えない）
                        # 2. $ifLinesForExpand 内の "varName.Append(...)" または "With varName"
                        $inferredVarName = $currentWithVar
                        if (-not $inferredVarName) {
                            foreach ($ifln in $ifLinesForExpand) {
                                # "varName.Append(...)" 形式
                                if ($ifln -match '(?i)^\s*(\w+)\.Append(?:Line)?\s*\(') {
                                    $inferredVarName = $Matches[1]
                                    break
                                }
                                # "With varName" 形式（With ブロック内の .Append を含むケース）
                                if ($ifln -match '(?i)^\s*With\s+(\w+)\s*$') {
                                    $inferredVarName = $Matches[1]
                                    break
                                }
                            }
                        }
                        if ($inferredVarName) {
                            if (-not $sbVars.ContainsKey($inferredVarName)) {
                                $sbVars[$inferredVarName] = @{
                                    Fragments  = [System.Collections.Generic.List[string]]::new()
                                    StartLine  = $lineNum
                                    EndLine    = $lineNum
                                    MethodName = $currentMethodName
                                }
                            }
                            foreach ($fragment in $branchResults) {
                                $sbVars[$inferredVarName].Fragments.Add($fragment)
                            }
                            $sbVars[$inferredVarName].EndLine = $endLineNum
                            $lastFragmentsList = $sbVars[$inferredVarName].Fragments
                            $lastVarName = $inferredVarName
                            $lastVarSource = $sbVars
                        }
                        else {
                            Write-Log -Level WARN -Message "Line ${lineNum}: If分岐の断片を関連付けるSQL変数が見つかりません" -LogFile $LogFile
                        }
                    }
                }
            }

            $i = $j
            continue
        }

        # ================================================
        # CommandText代入の検出
        # cmd.CommandText = "SELECT ..."
        # ================================================
        if ($trimmed -match '(?i)\.CommandText\s*=\s*(.+)$') {
            $cmdExpr = $Matches[1].Trim()
            # varName.ToString() → sbVarsの断片を確定してリセット（次のAppendが新しいSQLを開始できるよう）
            if ($cmdExpr -match '(?i)^(\w+)\.ToString\b') {
                $sbRefName = $Matches[1]
                if ($sbVars.ContainsKey($sbRefName) -and $sbVars[$sbRefName].Fragments.Count -gt 0) {
                    $prevInfo = $sbVars[$sbRefName]
                    $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                    $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                    if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                        $prevStmt = New-SqlStatement
                        $prevStmt.Sql = $prevMerged
                        $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                        $prevStmt.Category = 'Dynamic'
                        $prevStmt.StartLine = $prevInfo.StartLine
                        $prevStmt.EndLine = $lineNum
                        $prevStmt.SourceFile = $fileName
                        $prevStmt.MethodName = $currentMethodName
                        $sqlStatements.Add($prevStmt)
                    }
                    $sbVars.Remove($sbRefName)
                    $lastFragmentsList = $null
                    $lastVarName = $null
                    $lastVarSource = $null
                }
                continue
            }
            # 文字列リテラルを含まない場合は変数/メソッド参照のためスキップ
            if ($cmdExpr -notmatch '"') { continue }
            $sql = Extract-VbNetSqlFromExpression -Expression $cmdExpr

            if ($sql) {
                $stmt = New-SqlStatement
                $stmt.Sql = Convert-ToPlaceholder -SqlText $sql -Language 'vbnet'
                $stmt.Type = Get-SqlType -SqlText $sql
                $stmt.Category = 'Static'
                $stmt.StartLine = $lineNum
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $stmt.MethodName = $currentMethodName
                $sqlStatements.Add($stmt)
            }
            continue
        }

        # ================================================
        # New SqlCommand("SELECT ...", conn) の検出
        # ================================================
        if ($trimmed -match '(?i)New\s+(?:Sql|Oracle|OleDb|Odbc)?Command\s*\(\s*"(.+?)"') {
            $sql = $Matches[1] -replace '""', '"'

            $stmt = New-SqlStatement
            $stmt.Sql = Convert-ToPlaceholder -SqlText $sql -Language 'vbnet'
            $stmt.Type = Get-SqlType -SqlText $sql
            $stmt.Category = 'Static'
            $stmt.StartLine = $lineNum
            $stmt.EndLine = $lineNum
            $stmt.SourceFile = $fileName
            $stmt.MethodName = $currentMethodName
            $sqlStatements.Add($stmt)
            continue
        }

        # ================================================
        # StringBuilder.Append / .AppendLine の検出
        # ================================================
        if ($trimmed -match '(?i)^(\w+)\.Append(?:Line)?\s*\(\s*"(.+?)"\s*\)') {
            $sbVarName = $Matches[1]
            $sqlPart = $Matches[2] -replace '""', '"'

            # 別メソッドで同名変数が使われていた場合は既存断片を確定してリセット
            if ($sbVars.ContainsKey($sbVarName) -and $sbVars[$sbVarName].MethodName -ne $currentMethodName) {
                $prevInfo = $sbVars[$sbVarName]
                if ($prevInfo.Fragments.Count -gt 0) {
                    $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                    $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                    if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                        $prevStmt = New-SqlStatement
                        $prevStmt.Sql = $prevMerged
                        $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                        $prevStmt.Category = 'Dynamic'
                        $prevStmt.StartLine = $prevInfo.StartLine
                        $prevStmt.EndLine = $lineNum - 1
                        $prevStmt.SourceFile = $fileName
                        $prevStmt.MethodName = $prevInfo.MethodName
                        $sqlStatements.Add($prevStmt)
                    }
                }
                $sbVars.Remove($sbVarName)
                $lastFragmentsList = $null
                $lastVarName = $null
                $lastVarSource = $null
            }

            if (-not $sbVars.ContainsKey($sbVarName)) {
                $sbVars[$sbVarName] = @{
                    Fragments  = [System.Collections.Generic.List[string]]::new()
                    StartLine  = $lineNum
                    EndLine    = $lineNum
                    MethodName = $currentMethodName
                }
            }
            $sbVars[$sbVarName].Fragments.Add($sqlPart)
            $sbVars[$sbVarName].EndLine = $lineNum
            $lastFragmentsList = $sbVars[$sbVarName].Fragments
            $lastVarName = $sbVarName
            $lastVarSource = $sbVars
            continue
        }

        # ================================================
        # StringBuilder.Append にメソッド呼び出しが渡されている場合（部分抽出）
        # 例: sb.Append(BuildWithCteBlock("SELECT A, B, C FROM M_REF_TABLE "))
        # ================================================
        if ($trimmed -match '(?i)^(\w+)\.Append(?:Line)?\s*\(\s*([a-zA-Z_][\w.]*\s*\(.*\))\s*\)') {
            $sbVarName = $Matches[1]
            $callExpr  = $Matches[2].Trim()
            $placeholder = "/*:${callExpr}*/"

            # 別メソッドで同名変数が使われていた場合は既存断片を確定してリセット
            if ($sbVars.ContainsKey($sbVarName) -and $sbVars[$sbVarName].MethodName -ne $currentMethodName) {
                $prevInfo = $sbVars[$sbVarName]
                if ($prevInfo.Fragments.Count -gt 0) {
                    $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                    $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                    if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                        $prevStmt = New-SqlStatement
                        $prevStmt.Sql = $prevMerged
                        $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                        $prevStmt.Category = 'Dynamic'
                        $prevStmt.StartLine = $prevInfo.StartLine
                        $prevStmt.EndLine = $lineNum - 1
                        $prevStmt.SourceFile = $fileName
                        $prevStmt.MethodName = $prevInfo.MethodName
                        $sqlStatements.Add($prevStmt)
                    }
                }
                $sbVars.Remove($sbVarName)
                $lastFragmentsList = $null
                $lastVarName = $null
                $lastVarSource = $null
            }

            if (-not $sbVars.ContainsKey($sbVarName)) {
                $sbVars[$sbVarName] = @{
                    Fragments  = [System.Collections.Generic.List[string]]::new()
                    StartLine  = $lineNum
                    EndLine    = $lineNum
                    MethodName = $currentMethodName
                }
            }
            $sbVars[$sbVarName].Fragments.Add($placeholder)
            $sbVars[$sbVarName].EndLine = $lineNum
            $lastFragmentsList = $sbVars[$sbVarName].Fragments
            $lastVarName = $sbVarName
            $lastVarSource = $sbVars
            Write-Log -Level WARN -Message "Line ${lineNum}: メソッド呼び出し '$callExpr' を含むAppendを検出。SQL断片は不完全な可能性があります" -LogFile $LogFile
            continue
        }

        # ================================================
        # Dim varName As New StringBuilder の検出
        # 同名変数の再宣言時に既存断片を確定してリセットする
        # ================================================
        if ($trimmed -match '(?i)^Dim\s+(\w+)\s+As\s+New\s+(?:System\.Text\.)?StringBuilder\b') {
            $sbVarName = $Matches[1]
            if ($sbVars.ContainsKey($sbVarName) -and $sbVars[$sbVarName].Fragments.Count -gt 0) {
                $prevInfo = $sbVars[$sbVarName]
                $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                    $prevStmt = New-SqlStatement
                    $prevStmt.Sql = $prevMerged
                    $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                    $prevStmt.Category = 'Dynamic'
                    $prevStmt.StartLine = $prevInfo.StartLine
                    $prevStmt.EndLine = $lineNum - 1
                    $prevStmt.SourceFile = $fileName
                    $prevStmt.MethodName = $currentMethodName
                    $sqlStatements.Add($prevStmt)
                }
                $sbVars.Remove($sbVarName)
                # 最後に追跡した変数にかかわらず常にリセット（残存参照を確実にクリア）
                $lastFragmentsList = $null
                $lastVarName = $null
                $lastVarSource = $null
            }
            continue
        }

        # ================================================
        # sb = New StringBuilder() によるリセット検出
        # 同一メソッド内で同名変数を再生成して既存断片を確定する
        # ================================================
        if ($trimmed -match '(?i)^(\w+)\s*=\s*New\s+(?:System\.Text\.)?StringBuilder\b') {
            $sbVarName = $Matches[1]
            if ($sbVars.ContainsKey($sbVarName) -and $sbVars[$sbVarName].Fragments.Count -gt 0) {
                $prevInfo = $sbVars[$sbVarName]
                $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                    $prevStmt = New-SqlStatement
                    $prevStmt.Sql = $prevMerged
                    $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                    $prevStmt.Category = 'Dynamic'
                    $prevStmt.StartLine = $prevInfo.StartLine
                    $prevStmt.EndLine = $lineNum - 1
                    $prevStmt.SourceFile = $fileName
                    $prevStmt.MethodName = $currentMethodName
                    $sqlStatements.Add($prevStmt)
                }
                $sbVars.Remove($sbVarName)
                $lastFragmentsList = $null
                $lastVarName = $null
                $lastVarSource = $null
            }
            continue
        }

        # ================================================
        # sb.Clear() によるリセット検出
        # ================================================
        if ($trimmed -match '(?i)^(\w+)\.Clear\s*\(\s*\)\s*$') {
            $sbVarName = $Matches[1]
            if ($sbVars.ContainsKey($sbVarName) -and $sbVars[$sbVarName].Fragments.Count -gt 0) {
                $prevInfo = $sbVars[$sbVarName]
                $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                if ($prevMerged -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                    $prevStmt = New-SqlStatement
                    $prevStmt.Sql = $prevMerged
                    $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                    $prevStmt.Category = 'Dynamic'
                    $prevStmt.StartLine = $prevInfo.StartLine
                    $prevStmt.EndLine = $lineNum - 1
                    $prevStmt.SourceFile = $fileName
                    $prevStmt.MethodName = $currentMethodName
                    $sqlStatements.Add($prevStmt)
                }
                $sbVars.Remove($sbVarName)
                $lastFragmentsList = $null
                $lastVarName = $null
                $lastVarSource = $null
            }
            continue
        }

        # ================================================
        # Dim sql As String = "SELECT ..." の検出
        # ================================================
        if ($trimmed -match '(?i)^Dim\s+(\w+)\s+As\s+String\s*=\s*(.+)$') {
            $varName = $Matches[1]
            $assignExpr = $Matches[2].Trim()
            $sqlPart = Extract-VbNetSqlFromExpression -Expression $assignExpr

            if ($sqlPart -and $sqlPart -match '(?i)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                # 同名変数に既存断片がある場合は先に確定させる
                if ($dynamicSqlVars.ContainsKey($varName) -and $dynamicSqlVars[$varName].Fragments.Count -gt 0) {
                    $prevInfo = $dynamicSqlVars[$varName]
                    $prevMerged = Merge-DynamicSql -Fragments $prevInfo.Fragments.ToArray()
                    $prevMerged = Convert-ToPlaceholder -SqlText $prevMerged -Language 'vbnet'
                    $prevStmt = New-SqlStatement
                    $prevStmt.Sql = $prevMerged
                    $prevStmt.Type = Get-SqlType -SqlText $prevMerged
                    $prevStmt.Category = 'Dynamic'
                    $prevStmt.StartLine = $prevInfo.StartLine
                    $prevStmt.EndLine = $lineNum - 1
                    $prevStmt.SourceFile = $fileName
                    $prevStmt.MethodName = $currentMethodName
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
                $lastVarSource = $dynamicSqlVars
            }
            continue
        }

        # ================================================
        # sql &= "..." / sql += "..." の検出
        # ================================================
        if ($trimmed -match '(?i)^(\w+)\s*[&+]=\s*(.+)$') {
            $varName = $Matches[1]
            $appendExpr = $Matches[2].Trim()
            $sqlPart = Extract-VbNetSqlFromExpression -Expression $appendExpr

            if ($sqlPart) {
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
                $lastVarSource = $dynamicSqlVars
            }
            continue
        }

        # ================================================
        # sql = sql & "..." パターンの検出
        # Fix: -match は後方参照(\1)非対応のため [regex]::Match を使用
        # ================================================
        $m = [regex]::Match($trimmed, '(?i)^(\w+)\s*=\s*(\w+)\s*&\s*(.+)$')
        if ($m.Success -and $m.Groups[1].Value -ieq $m.Groups[2].Value) {
            $varName = $m.Groups[1].Value
            $appendExpr = $m.Groups[3].Value.Trim()
            $sqlPart = Extract-VbNetSqlFromExpression -Expression $appendExpr

            if ($sqlPart) {
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
                $lastVarSource = $dynamicSqlVars
            }
            continue
        }

        # ================================================
        # String.Format の検出
        # ================================================
        if ($trimmed -match '(?i)String\.Format\s*\(\s*"(.+?)"') {
            $formatSql = $Matches[1] -replace '""', '"'
            # {0}, {1} をプレースホルダに変換
            $formatSql = Convert-ToPlaceholder -SqlText $formatSql -Language 'vbnet'

            if ($formatSql -match '(?i)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                $stmt = New-SqlStatement
                $stmt.Sql = $formatSql
                $stmt.Type = Get-SqlType -SqlText $formatSql
                $stmt.Category = 'Dynamic'
                $stmt.StartLine = $lineNum
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $sqlStatements.Add($stmt)
            }
            continue
        }

        # ================================================
        # 補間文字列 $"SELECT ... {var} ..." の検出
        # ================================================
        if ($trimmed -match '(?i)\$"(.+?)"') {
            $interpSql = $Matches[1]
            $interpSql = Convert-ToPlaceholder -SqlText $interpSql -Language 'vbnet'

            if ($interpSql -match '(?i)^\s*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                $stmt = New-SqlStatement
                $stmt.Sql = $interpSql
                $stmt.Type = Get-SqlType -SqlText $interpSql
                $stmt.Category = 'Dynamic'
                $stmt.StartLine = $lineNum
                $stmt.EndLine = $lineNum
                $stmt.SourceFile = $fileName
                $sqlStatements.Add($stmt)
            }
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
            $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'

            # プレースホルダ /*:...*/ が先頭に来る場合も考慮してSQLキーワードを検出
            if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                $stmt = New-SqlStatement
                $stmt.Sql = $mergedSql
                $stmt.Type = Get-SqlType -SqlText $mergedSql
                $stmt.Category = 'Dynamic'
                $stmt.StartLine = $varInfo.StartLine
                $stmt.EndLine = $varInfo.EndLine
                $stmt.SourceFile = $fileName
                $stmt.MethodName = $currentMethodName
                $sqlStatements.Add($stmt)
            }
        }
    }

    # ================================================
    # StringBuilder変数をSQL文として確定
    # ================================================
    foreach ($sbEntry in $sbVars.GetEnumerator()) {
        $sbInfo = $sbEntry.Value
        if ($sbInfo.Fragments.Count -gt 0) {
            # Fragments には本文 + IF分岐断片が既に含まれている
            $mergedSql = Merge-DynamicSql -Fragments $sbInfo.Fragments.ToArray()
            $mergedSql = Convert-ToPlaceholder -SqlText $mergedSql -Language 'vbnet'

            # プレースホルダ /*:...*/ が先頭に来る場合も考慮してSQLキーワードを検出
            if ($mergedSql -match '(?i)^\s*(?:/\*.*?\*/\s*)*(SELECT|WITH|INSERT|UPDATE|DELETE|MERGE|CREATE|ALTER|DROP)\b') {
                $stmt = New-SqlStatement
                $stmt.Sql = $mergedSql
                $stmt.Type = Get-SqlType -SqlText $mergedSql
                $stmt.Category = 'Dynamic'
                $stmt.StartLine = $sbInfo.StartLine
                $stmt.EndLine = $sbInfo.EndLine
                $stmt.SourceFile = $fileName
                $stmt.MethodName = $currentMethodName
                $sqlStatements.Add($stmt)
            }
        }
    }

    $staticCount = ($sqlStatements | Where-Object { $_.Category -eq 'Static' }).Count
    $dynamicCount = ($sqlStatements | Where-Object { $_.Category -eq 'Dynamic' }).Count
    Write-Log -Level INFO -Message "Found $($sqlStatements.Count) SQL statements ($staticCount static, $dynamicCount dynamic)" -LogFile $LogFile

    return $sqlStatements.ToArray()
}

# ============================================================
# Join-VbNetContinuationLines: 行継続文字の結合
# 戻り値: @{Lines=[string[]]; OriginalLineNumbers=[int[]]}
#   Lines               - 行継続を結合した後の行配列
#   OriginalLineNumbers - 各行に対応する元ファイルの先頭行番号（1始まり）
# ============================================================
function Join-VbNetContinuationLines {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string[]]$Lines
    )

    $resultLines = [System.Collections.Generic.List[string]]::new()
    $resultNums  = [System.Collections.Generic.List[int]]::new()
    $buffer = ''
    $bufferStartLine = 0

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $trimmed = $Lines[$i].TrimEnd()

        if (-not $buffer) {
            $bufferStartLine = $i + 1  # 1始まり
        }

        if ($trimmed.EndsWith(' _') -or $trimmed.EndsWith("`t_")) {
            # 行継続: _ を除去して次行と結合
            $buffer += $trimmed.Substring(0, $trimmed.Length - 1).TrimEnd() + ' '
        }
        else {
            if ($buffer) {
                $resultLines.Add($buffer + $trimmed)
                $resultNums.Add($bufferStartLine)
                $buffer = ''
            }
            else {
                # $trimmed (TrimEnd済み) を使用し CRLF の \r を確実に除去する
                $resultLines.Add($trimmed)
                $resultNums.Add($i + 1)
            }
        }
    }

    # 最後のバッファが残っていれば追加
    if ($buffer) {
        $resultLines.Add($buffer.TrimEnd())
        $resultNums.Add($bufferStartLine)
    }

    return @{
        Lines               = $resultLines.ToArray()
        OriginalLineNumbers = $resultNums.ToArray()
    }
}

# ============================================================
# Remove-VbNetInlineComment: インラインコメントの除去
# ============================================================
function Remove-VbNetInlineComment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Line
    )

    # 文字列リテラル外の ' をコメントとして除去
    $inString = $false
    $result = [System.Text.StringBuilder]::new()

    for ($c = 0; $c -lt $Line.Length; $c++) {
        $ch = $Line[$c]

        if ($ch -eq '"') {
            $inString = -not $inString
            $result.Append($ch) | Out-Null
        }
        elseif ($ch -eq "'" -and -not $inString) {
            break  # コメント開始
        }
        else {
            $result.Append($ch) | Out-Null
        }
    }

    return $result.ToString().TrimEnd()
}

# ============================================================
# Extract-VbNetSqlFromExpression: VB.NET式からSQL部分を抽出
# ============================================================
function Extract-VbNetSqlFromExpression {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Expression
    )

    $expr = $Expression.Trim()

    # 連結演算子で分割
    $parts = $expr -split '\s*&\s*|\s*\+\s*'
    $fragments = [System.Collections.Generic.List[string]]::new()

    foreach ($part in $parts) {
        $p = $part.Trim()

        # 文字列リテラル: "..."
        if ($p -match '^"(.*)"$') {
            $literal = $Matches[1] -replace '""', '"'
            $fragments.Add($literal)
        }
        # vbCrLf / Environment.NewLine → 改行
        elseif ($p -match '(?i)^(vbCrLf|vbNewLine|Environment\.NewLine)$') {
            $fragments.Add("`n")
        }
        # vbTab → タブ
        elseif ($p -match '(?i)^vbTab$') {
            $fragments.Add("`t")
        }
        # 変数名 → プレースホルダ
        elseif ($p -match '^[a-zA-Z_][a-zA-Z0-9_.]*(?:\(.*\))?$') {
            # メソッド呼び出し（ToString等、括弧あり・なし両方）は除外
            if ($p -notmatch '(?i)\.(ToString|Trim|Replace|ToUpper|ToLower)\b') {
                $varName = $p -replace '\(.*\)', ''
                $fragments.Add("/*:$varName*/")
            }
        }
    }

    if ($fragments.Count -gt 0) {
        return ($fragments -join '')
    }
    return $null
}

# モジュールエクスポート
Export-ModuleMember -Function @(
    'Invoke-VbNetParser'
)
