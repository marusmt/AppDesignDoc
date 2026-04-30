$ErrorActionPreference = 'Stop'

# BFS で RootPath 配下を再帰走査し、言語別ファイル一覧を生成する
function Get-SourceFiles {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RootPath,
        [Parameter(Mandatory)][string[]]$ExcludeFolders,
        [Parameter(Mandatory)][string[]]$ExcludeExtensions,
        [Parameter(Mandatory)][string[]]$PlSqlExtensions,
        [Parameter(Mandatory)][string[]]$VbNetExtensions,
        [Parameter(Mandatory)][scriptblock]$LogCallback
    )

    # 末尾スラッシュを取り除いて正規化（R2）
    $normalizedRoot = $RootPath.TrimEnd('\', '/')

    # 拡張子を小文字でチェックするための hashtable を事前生成（R8）
    $excludeExtSet = @{}
    foreach ($e in $ExcludeExtensions) { $excludeExtSet[$e.ToLower()] = $true }
    $plSqlExtSet = @{}
    foreach ($e in $PlSqlExtensions)   { $plSqlExtSet[$e.ToLower()]   = $true }
    $vbNetExtSet = @{}
    foreach ($e in $VbNetExtensions)   { $vbNetExtSet[$e.ToLower()]   = $true }

    # PS 5.1 対応のキュー生成（R1: ::new() 不可）
    $queue = New-Object 'System.Collections.Generic.Queue[string]'
    $queue.Enqueue($normalizedRoot)

    $results = New-Object 'System.Collections.Generic.List[object]'
    $processedCount = 0

    while ($queue.Count -gt 0) {
        $currentDir = $queue.Dequeue()
        try {
            # サブフォルダを ExcludeFolders で前フィルタしてキューに追加
            $subDirs = Get-ChildItem -LiteralPath $currentDir -Directory -Force -ErrorAction Stop
            foreach ($subDir in $subDirs) {
                if ($ExcludeFolders -notcontains $subDir.Name) {
                    $queue.Enqueue($subDir.FullName)
                } else {
                    & $LogCallback 'SKIP' "スキップ（ExcludeFolders）: $($subDir.FullName)"
                }
            }

            # ファイルを走査して言語判定
            $files = Get-ChildItem -LiteralPath $currentDir -File -Force -ErrorAction Stop
            foreach ($file in $files) {
                $ext = [System.IO.Path]::GetExtension($file.Name).ToLower()
                $language =
                    if     ($excludeExtSet.ContainsKey($ext)) { 'SKIPPED' }
                    elseif ($plSqlExtSet.ContainsKey($ext))   { 'PLSQL'   }
                    elseif ($vbNetExtSet.ContainsKey($ext))   { 'VBNET'   }
                    else                                       { 'OTHER'   }

                if ($language -eq 'SKIPPED') {
                    & $LogCallback 'SKIP' "スキップ（ExcludeExtensions）: $($file.FullName)"
                }

                $relPath = $file.FullName.Substring($normalizedRoot.Length).TrimStart('\', '/')

                $results.Add([PSCustomObject]@{
                    FullPath      = $file.FullName
                    RelativePath  = $relPath
                    FileName      = $file.Name
                    Extension     = $ext
                    Language      = $language
                    SizeBytes     = $file.Length
                    LastWriteTime = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
                })

                $processedCount++
                # FR-09: 1,000 件ごとに進捗表示
                if ($processedCount % 1000 -eq 0) {
                    & $LogCallback 'INFO' "$processedCount 件処理済み"
                }
            }
        } catch [System.UnauthorizedAccessException] {
            # R5: アクセス拒否は握り潰して WARN ログのみ
            & $LogCallback 'WARN' "スキップ（アクセスエラー）: $currentDir"
        } catch {
            & $LogCallback 'WARN' "走査エラー: $currentDir : $($_.Exception.Message)"
        }
    }

    return $results.ToArray()
}

Export-ModuleMember -Function Get-SourceFiles
