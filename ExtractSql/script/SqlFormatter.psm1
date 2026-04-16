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

    # BETWEEN...AND を一時的に保護してから AND/OR の改行処理を行う
    $result = [regex]::Replace($result, '(?i)(BETWEEN\s+\S+\s+)AND\b', '$1__PROTECTED_AND__')

    # 主要句の前で改行+インデント
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

    # 連続する空行を除去
    $result = [regex]::Replace($result, '(\r?\n){3,}', "`n`n")

    # 末尾セミコロン付与
    $result = $result.TrimEnd()
    if (-not $result.EndsWith(';')) {
        $result += ';'
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

    $trimmed = $SqlText.TrimStart()

    switch -Regex ($trimmed) {
        '^(?i)SELECT'  { return 'SELECT' }
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
        [string]$Encoding = 'UTF8',

        [Parameter()]
        [string]$LogFile = ''
    )

    # 出力ディレクトリ作成
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
        Write-Log -Level INFO -Message "Created output directory: $OutputDir" -LogFile $LogFile
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($SourceFileName)
    $counter = 0
    $outputFiles = @()

    foreach ($stmt in $SqlStatements) {
        $counter++
        $seqNum = $counter.ToString('D3')
        $outputFileName = "${baseName}_${seqNum}.sql"
        $outputPath = Join-Path $OutputDir $outputFileName

        # ヘッダコメント生成
        $header = @"
-- ============================================
-- Source: $SourceFileName
-- Line: $($stmt.StartLine)-$($stmt.EndLine)
-- Type: $($stmt.Type) ($($stmt.Category))
-- Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
-- ============================================

"@

        # SQL文を整形
        $formattedSql = Format-SqlStatement -SqlText $stmt.Sql

        # ファイル書き込み
        $content = $header + $formattedSql
        $content | Out-File -FilePath $outputPath -Encoding $Encoding -Force

        $outputFiles += $outputPath
        Write-Log -Level INFO -Message "Output: $outputPath" -LogFile $LogFile
    }

    return $outputFiles
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

    foreach ($line in $Lines) {
        # IF開始
        if ($line -match $ifPattern -and $nestLevel -eq 0) {
            $branchCount++
            $currentCondition = $Matches[1]
            $result.Add("-- [Branch $branchCount] $currentCondition")
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
            $result.Add("-- [Branch $branchCount] $currentCondition")
            continue
        }

        # ELSE
        if ($line -match $elsePattern -and $nestLevel -eq 1) {
            $branchCount++
            $result.Add("-- [Branch $branchCount] ELSE")
            continue
        }

        # END IF（トップレベル）
        if ($line -match $endIfPattern -and $nestLevel -eq 1) {
            $nestLevel = 0
            $inBranch = $false
            continue
        }

        # 分岐内のSQL断片を抽出
        if ($inBranch -or $nestLevel -gt 0) {
            $sqlFragment = & $ExtractSqlFromLine $line
            if ($sqlFragment) {
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
    'Write-Log',
    'Expand-IfBranches'
)
