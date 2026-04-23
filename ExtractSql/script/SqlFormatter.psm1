#Requires -Version 5.1
<#
.SYNOPSIS
    SQL整形・プレースホルダ変換・ファイル出力モジュール
.DESCRIPTION
    抽出されたSQL文を整形し、変数をプレースホルダに変換し、
    単独実行可能な .sql ファイルとして出力します。
#>

# ============================================================
# SQL文情報を格納するクラス
# ============================================================
class SqlStatement {
    [string]$Sql
    [string]$Type          # SELECT / INSERT / UPDATE / DELETE / MERGE / DDL / OTHER
    [string]$Category      # Static / Dynamic
    [int]$StartLine
    [int]$EndLine
    [string]$SourceFile
    [string]$MethodName    # 抽出元のメソッド名（Sub/Function）
    [string]$CursorName    # CURSOR 定義名（PL/SQL）
    [System.Collections.Generic.List[string]]$BranchComments

    SqlStatement() {
        $this.BranchComments = [System.Collections.Generic.List[string]]::new()
    }
}

# ============================================================
# New-SqlStatement: SqlStatementオブジェクトのファクトリ関数
# using module による型の二重ロードを回避するため、
# パーサーモジュールはこの関数経由でオブジェクトを生成する
# ============================================================
function New-SqlStatement {
    return [SqlStatement]::new()
}

# ============================================================
# Merge-DynamicSql: 動的SQL文字列の結合
# ============================================================
function Merge-DynamicSql {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Fragments
    )

    $merged = ($Fragments -join '') -replace '\s+', ' '
    return $merged.Trim()
}

# ============================================================
# Convert-ToPlaceholder: 変数→プレースホルダ変換
# ============================================================
function Convert-ToPlaceholder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SqlText,

        [Parameter()]
        [string]$Language = 'plsql'  # plsql / vbnet
    )

    $result = $SqlText

    if ($Language -eq 'plsql') {
        # PL/SQL: || 演算子による連結の変数部分をプレースホルダに変換
        # パターン: || variable_name || or || variable_name;
        # 既に Merge 済みの場合は変数名が残っているケースを処理
        $result = [regex]::Replace($result,
            '\|\|\s*([a-zA-Z_][a-zA-Z0-9_.]*?)\s*(?:\|\||$|;)',
            '/*:$1*/ ')
    }
    elseif ($Language -eq 'vbnet') {
        # VB.NET: String.Format の {0}, {1} をプレースホルダに変換
        $result = [regex]::Replace($result,
            '\{(\d+)\}',
            '/*:param$1*/')

        # VB.NET: 補間文字列 {varName} をプレースホルダに変換
        $result = [regex]::Replace($result,
            '\{([a-zA-Z_][a-zA-Z0-9_.]*)\}',
            '/*:$1*/')
    }

    # 連続する空白を正規化
    $result = $result -replace '\s{2,}', ' '

    return $result.Trim()
}

# ============================================================
# Format-SqlStatement: SQL文の整形
# ============================================================
function Format-SqlStatement {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SqlText
    )

    $result = $SqlText.Trim()

    # メソッド呼び出しプレースホルダ /*:method(...)*/  を -- [未展開: ...] コメント行に変換（SQL先頭のみ）
    # 変数プレースホルダ /*:varName*/ はそのまま保持する
    $prefixComments = [System.Collections.Generic.List[string]]::new()
    while ($result.StartsWith('/*:')) {
        $closingIdx = $result.IndexOf('*/')
        if ($closingIdx -lt 0) { break }
        $content = $result.Substring(3, $closingIdx - 3).Trim()
        # メソッド呼び出し形式（識別子 + 括弧）の場合のみ変換
        if ($content -match '^[a-zA-Z_][\w.]*\s*\(') {
            $prefixComments.Add("-- [未展開: ${content}]")
            $result = $result.Substring($closingIdx + 2).TrimStart()
        } else {
            break
        }
    }

    # SQLキーワードを大文字に統一
    $keywords = @(
        'SELECT', 'FROM', 'WHERE', 'AND', 'OR', 'NOT',
        'INSERT', 'INTO', 'VALUES', 'UPDATE', 'SET',
        'DELETE', 'MERGE', 'USING', 'WHEN', 'MATCHED', 'THEN',
        'JOIN', 'INNER', 'LEFT', 'RIGHT', 'OUTER', 'CROSS', 'FULL',
        'ON', 'AS', 'IN', 'EXISTS', 'BETWEEN', 'LIKE', 'IS', 'NULL',
        'ORDER', 'BY', 'GROUP', 'HAVING', 'UNION', 'ALL', 'DISTINCT',
        'CREATE', 'ALTER', 'DROP', 'TABLE', 'INDEX', 'VIEW',
        'BEGIN', 'END', 'DECLARE', 'CURSOR', 'OPEN', 'FETCH', 'CLOSE',
        'CASE', 'ELSE', 'WHEN', 'END',
        'COUNT', 'SUM', 'AVG', 'MIN', 'MAX',
        'ASC', 'DESC', 'LIMIT', 'OFFSET', 'ROWNUM',
        'WITH', 'RECURSIVE', 'OVER', 'PARTITION',
        'GRANT', 'REVOKE', 'COMMIT', 'ROLLBACK',
        'TRUNCATE', 'EXECUTE', 'IMMEDIATE'
    )

    foreach ($kw in $keywords) {
        # 単語境界で一致するキーワードを大文字に変換
        $pattern = '(?i)\b' + [regex]::Escape($kw) + '\b'
        $result = [regex]::Replace($result, $pattern, $kw)
    }

    # 改行がない場合のみ句キーワードの前で改行+インデントを付与する
    # （改行が既にある場合は元のソース形式を保持）
    if ($result -notmatch "`n") {
        # BETWEEN...AND を一時的に保護してから AND/OR の改行処理を行う
        $result = [regex]::Replace($result, '(?i)(BETWEEN\s+\S+\s+)AND\b', '$1__PROTECTED_AND__')

        $clauseKeywords = @('FROM', 'WHERE', 'AND', 'OR', 'ORDER BY',
            'GROUP BY', 'HAVING', 'JOIN', 'INNER JOIN', 'LEFT JOIN',
            'RIGHT JOIN', 'FULL JOIN', 'CROSS JOIN', 'ON',
            'SET', 'VALUES', 'INTO', 'USING', 'WHEN MATCHED', 'UNION')

        foreach ($clause in $clauseKeywords) {
            $pattern = '(?<!--)\s+(?=' + [regex]::Escape($clause) + '\b)'
            $result = [regex]::Replace($result, $pattern, "`n  ")
        }

        # BETWEEN...AND を復元
        $result = $result -replace '__PROTECTED_AND__', 'AND'
    }

    # 分岐コメント /* [Branch N] ... */ の前後に改行を付与して視覚的に分離する
    # 句キーワード改行の後に処理することで FROM/WHERE 等の改行も正常に機能する
    # /*:varName*/ プレースホルダとは区別するため [Branch を明示的に検索する
    $result = [regex]::Replace($result, '\s*(/\*\s*\[Branch\s+\d+\][^*]*\*/)\s*', "`n`$1`n  ")

    # 空行を除去
    $result = ($result -split '\r?\n' | Where-Object { $_.Trim() -ne '' }) -join "`n"

    # 末尾セミコロン付与
    $result = $result.TrimEnd()
    if (-not $result.EndsWith(';')) {
        $result += ';'
    }

    # 先頭のメソッド呼び出しコメント行を付与
    if ($prefixComments.Count -gt 0) {
        $result = ($prefixComments -join "`n") + "`n" + $result
    }

    return $result
}

# ============================================================
# Get-SqlType: SQL文の種別判定
# ============================================================
function Get-SqlType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SqlText
    )

    # プレースホルダ /*:...*/ が先頭に来る場合は除去してからキーワード判定
    $trimmed = ($SqlText.TrimStart() -replace '^(?:/\*.*?\*/\s*)+', '').TrimStart()

    switch -Regex ($trimmed) {
        '^(?i)SELECT'  { return 'SELECT' }
        '^(?i)WITH'    { return 'SELECT' }
        '^(?i)INSERT'  { return 'INSERT' }
        '^(?i)UPDATE'  { return 'UPDATE' }
        '^(?i)DELETE'  { return 'DELETE' }
        '^(?i)MERGE'   { return 'MERGE' }
        '^(?i)CREATE'  { return 'DDL' }
        '^(?i)ALTER'   { return 'DDL' }
        '^(?i)DROP'    { return 'DDL' }
        '^(?i)TRUNCATE' { return 'DDL' }
        '^(?i)GRANT'   { return 'DCL' }
        '^(?i)REVOKE'  { return 'DCL' }
        default        { return 'OTHER' }
    }
}

# ============================================================
# Export-SqlFiles: SQLファイルの出力
# ============================================================
function Export-SqlFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$SqlStatements,

        [Parameter(Mandatory)]
        [string]$SourceFileName,

        [Parameter(Mandatory)]
        [string]$OutputDir,

        [Parameter()]
        [string]$Encoding = 'Default',

        [Parameter()]
        [ValidateSet('PerSql', 'PerSource')]
        [string]$OutputFormat = 'PerSql',

        [Parameter()]
        [string]$LogFile = ''
    )

    # 出力ディレクトリ作成
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        Write-Log -Level INFO -Message "Created output directory: $OutputDir" -LogFile $LogFile
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceFileName)
    $outputFiles = @()
    $totalCount = $SqlStatements.Count

    if ($OutputFormat -eq 'PerSource') {
        # ----------------------------------------
        # PerSource: ソースファイル単位で1ファイルに出力
        # ----------------------------------------
        $outputFileName = "${baseName}.sql"
        $outputPath = Join-Path $OutputDir $outputFileName
        $contentParts = [System.Collections.Generic.List[string]]::new()

        $counter = 0
        foreach ($stmt in $SqlStatements) {
            $counter++
            $methodLine = if ($stmt.MethodName) { "`n-- Method: $($stmt.MethodName)" } else { '' }
            $cursorLine = if ($stmt.CursorName) { "`n-- Cursor: $($stmt.CursorName)" } else { '' }
            $header = @"
-- ============================================
-- Source: $SourceFileName$methodLine$cursorLine
-- SQL: $counter / $totalCount
-- Line: $($stmt.StartLine)-$($stmt.EndLine)
-- Type: $($stmt.Type) ($($stmt.Category))
-- Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
-- ============================================

"@
            $formattedSql = Format-SqlStatement -SqlText $stmt.Sql
            $contentParts.Add($header + $formattedSql)
        }

        $content = $contentParts -join "`n`n"
        $content | Out-File -FilePath $outputPath -Encoding $Encoding -Force

        $outputFiles += $outputPath
        Write-Log -Level INFO -Message "Output: $outputPath ($totalCount SQLs)" -LogFile $LogFile
    }
    else {
        # ----------------------------------------
        # PerSql: SQL毎に個別ファイルに出力（既存動作）
        # ----------------------------------------
        $counter = 0
        foreach ($stmt in $SqlStatements) {
            $counter++
            $seqNum = $counter.ToString('D3')
            $outputFileName = "${baseName}_${seqNum}.sql"
            $outputPath = Join-Path $OutputDir $outputFileName

            $methodLine = if ($stmt.MethodName) { "`n-- Method: $($stmt.MethodName)" } else { '' }
            $cursorLine = if ($stmt.CursorName) { "`n-- Cursor: $($stmt.CursorName)" } else { '' }
            $header = @"
-- ============================================
-- Source: $SourceFileName$methodLine$cursorLine
-- Line: $($stmt.StartLine)-$($stmt.EndLine)
-- Type: $($stmt.Type) ($($stmt.Category))
-- Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
-- ============================================

"@
            $formattedSql = Format-SqlStatement -SqlText $stmt.Sql

            $content = $header + $formattedSql
            $content | Out-File -FilePath $outputPath -Encoding $Encoding -Force

            $outputFiles += $outputPath
            Write-Log -Level INFO -Message "Output: $outputPath" -LogFile $LogFile
        }
    }

    return $outputFiles
}

# ============================================================
# Get-FileEncoding: BOMによるファイルエンコーディング自動検出
# BOMが検出された場合はそのエンコーディングを返す。
# BOMがない場合は FallbackEncoding を返す。
# ============================================================
function Get-FileEncoding {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string]$FallbackEncoding = 'Default'
    )

    $bom = New-Object byte[] 4
    $stream = [System.IO.File]::OpenRead($FilePath)
    try {
        $read = $stream.Read($bom, 0, 4)
    }
    finally {
        $stream.Close()
    }

    if ($read -ge 3 -and $bom[0] -eq 0xEF -and $bom[1] -eq 0xBB -and $bom[2] -eq 0xBF) {
        return 'UTF8'             # UTF-8 with BOM
    }
    elseif ($read -ge 2 -and $bom[0] -eq 0xFF -and $bom[1] -eq 0xFE) {
        return 'Unicode'          # UTF-16 LE
    }
    elseif ($read -ge 2 -and $bom[0] -eq 0xFE -and $bom[1] -eq 0xFF) {
        return 'BigEndianUnicode' # UTF-16 BE
    }
    else {
        return $FallbackEncoding  # BOMなし → フォールバック
    }
}

# ============================================================
# Write-Log: ログ出力
# ============================================================
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [string]$LogFile = ''
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logEntry = "[$Level]  $Message"
    $logEntryWithTime = "$timestamp $logEntry"

    # コンソール出力（色分け）
    switch ($Level) {
        'INFO'  { Write-Host $logEntry -ForegroundColor Cyan }
        'WARN'  { Write-Host $logEntry -ForegroundColor Yellow }
        'ERROR' { Write-Host $logEntry -ForegroundColor Red }
    }

    # ログファイル出力
    if ($LogFile) {
        $logEntryWithTime | Out-File -FilePath $LogFile -Append -Encoding UTF8
    }
}

# ============================================================
# Expand-IfBranches: IF分岐の展開（共通ロジック）
# ============================================================
function Expand-IfBranches {
    <#
    .SYNOPSIS
        IF分岐内のSQL断片をすべて展開し、分岐コメント付きで返す
    .DESCRIPTION
        分岐構造を解析し、制御構文を除去して
        すべての分岐パスのSQL断片をフラットに展開します。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string[]]$Lines,

        [Parameter(Mandatory)]
        [string]$Language,  # plsql / vbnet

        [Parameter(Mandatory)]
        [scriptblock]$ExtractSqlFromLine  # 各行からSQL断片を抽出するロジック
    )

    $result = [System.Collections.Generic.List[string]]::new()
    $branchCount = 0
    $nestLevel = 0
    $inBranch = $false
    $currentCondition = ''
    $pendingBranchComment = $null  # SQL断片が見つかった時のみ出力する分岐コメントバッファ

    # 言語別の正規表現パターン
    if ($Language -eq 'plsql') {
        $ifPattern      = '(?i)^\s*IF\s+(.+?)\s+THEN\s*$'
        $elsifPattern   = '(?i)^\s*ELSIF\s+(.+?)\s+THEN\s*$'
        $elsePattern    = '(?i)^\s*ELSE\s*$'
        $endIfPattern   = '(?i)^\s*END\s+IF\s*;?\s*$'
        $nestedIfPattern = '(?i)^\s*IF\s+'
    }
    else {
        # VB.NET
        $ifPattern      = '(?i)^\s*If\s+(.+?)\s+Then\s*$'
        $elsifPattern   = '(?i)^\s*ElseIf\s+(.+?)\s+Then\s*$'
        $elsePattern    = '(?i)^\s*Else\s*$'
        $endIfPattern   = '(?i)^\s*End\s+If\s*$'
        $nestedIfPattern = '(?i)^\s*If\s+'
    }

    foreach ($rawLine in $Lines) {
        # Windows CRLF 対策: \r を除去して正規化
        $line = $rawLine.TrimEnd("`r")
        # VB.NET コメント行のスキップ（' で始まる行）
        if ($Language -eq 'vbnet' -and $line.Trim() -match "^\s*'") {
            continue
        }
        # IF開始
        if ($line -match $ifPattern -and $nestLevel -eq 0) {
            $branchCount++
            $currentCondition = $Matches[1]
            $pendingBranchComment = "/* [Branch $branchCount] $currentCondition */"
            $inBranch = $true
            $nestLevel = 1
            continue
        }

        # ネストされたIF
        if ($line -match $nestedIfPattern -and $nestLevel -gt 0) {
            $nestLevel++
            # ネスト内もSQL断片を抽出して展開（以降の行処理に続く）
        }

        # ネストされたEND IF
        if ($line -match $endIfPattern -and $nestLevel -gt 1) {
            $nestLevel--
            continue
        }

        # ELSIF
        if ($line -match $elsifPattern -and $nestLevel -eq 1) {
            $branchCount++
            $currentCondition = $Matches[1]
            $pendingBranchComment = "/* [Branch $branchCount] $currentCondition */"
            continue
        }

        # ELSE
        if ($line -match $elsePattern -and $nestLevel -eq 1) {
            $branchCount++
            $pendingBranchComment = "/* [Branch $branchCount] ELSE */"
            continue
        }

        # END IF（トップレベル）
        if ($line -match $endIfPattern -and $nestLevel -eq 1) {
            $nestLevel = 0
            $inBranch = $false
            $pendingBranchComment = $null
            continue
        }

        # 分岐内のSQL断片を抽出
        if ($inBranch -or $nestLevel -gt 0) {
            # VB.NET: sb = New StringBuilder → SQL の区切りをセンチネルで通知
            if ($Language -eq 'vbnet' -and $line -match '(?i)^\s*(\w+)\s*=\s*New\s+(?:System\.Text\.)?StringBuilder\b') {
                $result.Add('$$SQL_RESET:' + $Matches[1] + '$$')
                continue
            }
            $sqlFragment = & $ExtractSqlFromLine $line
            if ($sqlFragment) {
                if ($pendingBranchComment) {
                    $result.Add($pendingBranchComment)
                    $pendingBranchComment = $null
                }
                $result.Add($sqlFragment)
            }
        }
    }

    return $result.ToArray()
}

# モジュールエクスポート
Export-ModuleMember -Function @(
    'New-SqlStatement',
    'Merge-DynamicSql',
    'Convert-ToPlaceholder',
    'Format-SqlStatement',
    'Get-SqlType',
    'Export-SqlFiles',
    'Get-FileEncoding',
    'Write-Log',
    'Expand-IfBranches'
)
