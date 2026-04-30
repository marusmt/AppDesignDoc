$ErrorActionPreference = 'Stop'

$Script:LogWriteFailed = $false

function Write-LogMessage {
    <#
    .SYNOPSIS
        指定されたレベルとメッセージをコンソールおよびログファイルに出力する。
    .DESCRIPTION
        タイムスタンプ・レベル・メッセージの形式でコンソール出力（Write-Host）と
        ファイル追記（Add-Content -Encoding UTF8）を行う。
        ログファイルへの書き込みに失敗した場合、モジュールスコープの失敗フラグを立て、
        以降の呼び出しではファイル書き込みを試行せずコンソール出力のみ継続する。
    .PARAMETER Level
        ログレベル。'INFO' / 'WARN' / 'SKIP' / 'ERROR' のいずれか。
    .PARAMETER Message
        ログ本文。
    .PARAMETER LogPath
        ログファイルの絶対パス。呼び出し側が事前にディレクトリを作成しておくこと。
    .EXAMPLE
        Write-LogMessage -Level 'INFO' -Message '走査を開始します' -LogPath '.\logs\SourceList_20260429.log'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'SKIP', 'ERROR')]
        [string]$Level,
        [Parameter(Mandatory)]
        [string]$Message,
        [Parameter(Mandatory)]
        [string]$LogPath
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$($Level.PadRight(5))] $Message"

    Write-Host $line

    if (-not $Script:LogWriteFailed) {
        try {
            Add-Content -Path $LogPath -Value $line -Encoding UTF8
        } catch {
            $Script:LogWriteFailed = $true
            Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [WARN ] ログファイルへの書き込みに失敗しました: $($_.Exception.Message)"
        }
    }
}

Export-ModuleMember -Function Write-LogMessage
