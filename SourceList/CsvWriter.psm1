$ErrorActionPreference = 'Stop'

function ConvertTo-CsvField {
    param([string]$Value)
    '"' + $Value.Replace('"', '""') + '"'
}

function Export-SourceListCsv {
    <#
    .SYNOPSIS
        走査結果レコードを UTF-8 BOM 付き CSV ファイルとして出力する。
    .PARAMETER Records
        Get-SourceFiles + Add-Member 後のファイルオブジェクト配列。
    .PARAMETER OutputPath
        出力先 CSV ファイルの絶対パス（SourceList.ps1 側でタイムスタンプ付きパスを決定）。
    .PARAMETER LogCallback
        ログ出力用コールバック { param($Level, $Message) }。省略可。
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Records,
        [Parameter(Mandatory)]
        [string]$OutputPath,
        [scriptblock]$LogCallback
    )

    $header = '"FullPath","RelativePath","FileName","Extension","Language",' +
              '"SizeBytes","Lines","CommentLines","BlankLines","LastWriteTime"'

    $writer = New-Object System.IO.StreamWriter($OutputPath, $false, [System.Text.Encoding]::UTF8)
    $writer.NewLine = "`r`n"
    try {
        $writer.WriteLine($header)

        foreach ($rec in $Records) {
            $row = @(
                (ConvertTo-CsvField ([string]$rec.FullPath)),
                (ConvertTo-CsvField ([string]$rec.RelativePath)),
                (ConvertTo-CsvField ([string]$rec.FileName)),
                (ConvertTo-CsvField ([string]$rec.Extension)),
                (ConvertTo-CsvField ([string]$rec.Language)),
                (ConvertTo-CsvField ([string]$rec.SizeBytes)),
                (ConvertTo-CsvField ([string]$rec.Lines)),
                (ConvertTo-CsvField ([string]$rec.CommentLines)),
                (ConvertTo-CsvField ([string]$rec.BlankLines)),
                (ConvertTo-CsvField ([string]$rec.LastWriteTime))
            ) -join ','
            $writer.WriteLine($row)
        }

        if ($null -ne $LogCallback) {
            & $LogCallback 'INFO' "CSV 出力完了: $OutputPath ($($Records.Count) 件)"
        }
    } finally {
        $writer.Close()
    }
}

Export-ModuleMember -Function Export-SourceListCsv