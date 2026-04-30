$ErrorActionPreference = 'Stop'

function Get-FileEncoding {
    <#
    .SYNOPSIS
        ファイルの文字エンコーディングを BOM 検出 + 拡張子フォールバックで判定する。
    .PARAMETER FilePath
        判定対象ファイルの絶対パス。
    .PARAMETER DefaultEncoding
        BOM なし・非 PL/SQL 拡張子の場合に返す値。既定は 'UTF-8'。
    .OUTPUTS
        文字列: 'UTF-8-BOM' | 'UTF-16LE' | 'UTF-16BE' | 'SHIFT-JIS' | 'UTF-8'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,
        [string]$DefaultEncoding = 'UTF-8'
    )

    # ① 先頭 4 バイトを読む
    $fs = $null
    $buf = New-Object byte[] 4
    $bytesRead = 0
    try {
        $fs = [System.IO.File]::OpenRead($FilePath)
        $bytesRead = $fs.Read($buf, 0, 4)
    } catch {
        return $DefaultEncoding
    } finally {
        if ($null -ne $fs) { $fs.Dispose() }
    }

    # ② BOM 判定（優先順位: UTF-8 BOM > UTF-16 LE > UTF-16 BE）
    if ($bytesRead -ge 3 -and $buf[0] -eq 0xEF -and $buf[1] -eq 0xBB -and $buf[2] -eq 0xBF) {
        return 'UTF-8-BOM'
    }
    if ($bytesRead -ge 2 -and $buf[0] -eq 0xFF -and $buf[1] -eq 0xFE) {
        return 'UTF-16LE'
    }
    if ($bytesRead -ge 2 -and $buf[0] -eq 0xFE -and $buf[1] -eq 0xFF) {
        return 'UTF-16BE'
    }

    # ③ BOM なし → 拡張子フォールバック
    $ext = [System.IO.Path]::GetExtension($FilePath).ToLower()
    $plsqlExts = @('.sql', '.pkb', '.pks', '.prc', '.fnc', '.trg', '.vw')
    if ($plsqlExts -contains $ext) {
        return 'SHIFT-JIS'
    }

    return $DefaultEncoding
}

Export-ModuleMember -Function Get-FileEncoding
