$ErrorActionPreference = 'Stop'

function Get-LineCounts {
    <#
    .SYNOPSIS
        ファイルの総行数・コメント行数・空行数をカウントする。
    .PARAMETER FilePath
        カウント対象ファイルの絶対パス。
    .PARAMETER Language
        言語種別。'VBNET' / 'PLSQL' / 'OTHER' / 'SKIPPED' のいずれか。
    .PARAMETER Encoding
        Get-FileEncoding が返す文字列。
    .OUTPUTS
        PSCustomObject: { Lines: int; CommentLines: int; BlankLines: int }
        Lines = -1, BlankLines = -1 は読み取りエラーを示す。
    .NOTES
        [VB.NET 制限] 文字列リテラル内の ' による誤検出は実用上発生しない（行頭に来ないため）。
        [PL/SQL 制限] 文字列リテラル内の '--'/'/*' の誤検出あり（行頭 '--' は実用上ほぼ発生しない）。
        [PL/SQL 制限] ネストした /* /* */ */ は未対応（PL/SQL 標準外）。
        [PL/SQL 制限] 同一行の複数 /* */ ペアは最初のもののみ考慮。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        [Parameter(Mandatory)]
        [string]$Language,
        [Parameter(Mandatory)]
        [string]$Encoding
    )

    if ($Language -eq 'SKIPPED') {
        return [PSCustomObject]@{ Lines = 0; CommentLines = 0; BlankLines = 0 }
    }

    $enc = switch ($Encoding) {
        'UTF-8-BOM' { [System.Text.Encoding]::UTF8 }
        'UTF-16LE'  { [System.Text.Encoding]::Unicode }
        'UTF-16BE'  { [System.Text.Encoding]::BigEndianUnicode }
        'SHIFT-JIS' { [System.Text.Encoding]::GetEncoding(932) }
        default     { New-Object System.Text.UTF8Encoding($false) }
    }

    try {
        $allLines = [System.IO.File]::ReadAllLines($FilePath, $enc)
    } catch {
        return [PSCustomObject]@{ Lines = -1; CommentLines = 0; BlankLines = -1 }
    }

    $blankLines = 0
    foreach ($line in $allLines) {
        if ($line.Trim() -eq '') { $blankLines++ }
    }

    $commentLines = 0

    # ── VB.NET コメント判定 ─────────────────────────────────────────────────
    if ($Language -eq 'VBNET') {
        foreach ($line in $allLines) {
            $trimmed = $line.TrimStart()
            if ($trimmed.Length -eq 0) { continue }

            if ($trimmed[0] -eq "'") {
                $commentLines++
            } elseif ($trimmed.Length -ge 3) {
                $rem3 = $trimmed.Substring(0, 3)
                if ($rem3 -ieq 'REM' -and (
                    $trimmed.Length -eq 3 -or
                    $trimmed[3] -eq ' '   -or
                    $trimmed[3] -eq "`t"
                )) {
                    $commentLines++
                }
            }
        }
    }

    # ── PL/SQL コメント判定（3 状態機械: Normal / InBlock / InHint）────────
    if ($Language -eq 'PLSQL') {
        $inBlock = $false
        $inHint  = $false

        foreach ($line in $allLines) {
            $trimmed = $line.TrimStart()

            if ($inBlock) {
                $closeIdx = $line.IndexOf('*/')
                if ($closeIdx -ge 0) {
                    $inBlock = $false
                    if ($line.Substring($closeIdx + 2).Trim() -eq '') { $commentLines++ }
                } else {
                    $commentLines++
                }

            } elseif ($inHint) {
                $closeIdx = $line.IndexOf('*/')
                if ($closeIdx -ge 0) { $inHint = $false }

            } else {
                if ($trimmed.Length -ge 2 -and $trimmed[0] -eq '-' -and $trimmed[1] -eq '-') {
                    $commentLines++
                } else {
                    $openIdx = $line.IndexOf('/*')
                    if ($openIdx -ge 0) {
                        $isHint     = ($openIdx + 2 -lt $line.Length) -and ($line[$openIdx + 2] -eq '+')
                        $searchFrom = if ($isHint) { $openIdx + 3 } else { $openIdx + 2 }
                        $closeIdx   = if ($searchFrom -lt $line.Length) {
                                          $line.IndexOf('*/', $searchFrom)
                                      } else { -1 }
                        $beforeOpen = $line.Substring(0, $openIdx).Trim()
                        $afterClose = if ($closeIdx -ge 0) { $line.Substring($closeIdx + 2).Trim() } else { $null }

                        if ($isHint) {
                            if ($closeIdx -lt 0) { $inHint = $true }
                        } else {
                            if ($beforeOpen -eq '') {
                                if ($closeIdx -ge 0) {
                                    if ($afterClose -eq '') { $commentLines++ }
                                } else {
                                    $inBlock = $true
                                    $commentLines++
                                }
                            } else {
                                if ($closeIdx -lt 0) { $inBlock = $true }
                            }
                        }
                    }
                }
            }
        }
    }

    return [PSCustomObject]@{
        Lines        = $allLines.Count
        CommentLines = $commentLines
        BlankLines   = $blankLines
    }
}

Export-ModuleMember -Function Get-LineCounts