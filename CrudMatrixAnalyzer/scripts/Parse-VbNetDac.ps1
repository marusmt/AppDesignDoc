<#
.SYNOPSIS
    VB.NET DACファイルを解析し、CRUD操作情報を抽出する

.DESCRIPTION
    VB.NETのデータアクセス層（DAC）ソースファイルから埋め込みSQL文を抽出し、
    INSERT/SELECT/UPDATE/DELETE/MERGE文のテーブル名・項目名・操作種別を解析する
    対象: ファイル名に "dac" を含む .vb ファイル

.PARAMETER SourcePath
    VB.NETソースファイル格納ディレクトリ

.PARAMETER DacFilePattern
    DACファイルのパターン（デフォルト: *dac*.vb）
#>

function Get-VbNetSqlStrings {
    param([string]$Content)

    $Content = $Content -replace [char]0x3000, ' '

    $sqlStrings = [System.Collections.ArrayList]::new()

    # パターン1: 単一行の文字列リテラル（SQL キーワードを含むもの）
    $singleLinePattern = '"([^"]*(?:SELECT|INSERT|UPDATE|DELETE|MERGE)[^"]*)"'
    $singleMatches = [regex]::Matches($Content, $singleLinePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    foreach ($match in $singleMatches) {
        $sqlText = $match.Groups[1].Value
        [void]$sqlStrings.Add($sqlText)
    }

    # パターン2: 文字列連結（& または +）で複数行にまたがるSQL
    $concatPattern = '(?i)(?:Dim|Const)\s+\w+\s+As\s+String\s*=\s*"([^"]*)"(?:\s*(?:&|_\s*\r?\n\s*&?|\+)\s*"([^"]*)")*'
    $concatMatches = [regex]::Matches($Content, $concatPattern)

    foreach ($match in $concatMatches) {
        $segments = [regex]::Matches($match.Value, '"([^"]*)"')
        if ($segments.Count -le 1) { continue }
        $fullSql = ($segments | ForEach-Object { $_.Groups[1].Value }) -join " "
        if ($fullSql -match '(?i)(SELECT|INSERT|UPDATE|DELETE|MERGE)') {
            [void]$sqlStrings.Add($fullSql)
        }
    }

    # パターン3: StringBuilder.Append / AppendLine パターン
    $sbPattern = '(?i)(?:Dim|Const)\s+(\w+)\s+As\s+(?:New\s+)?(?:System\.Text\.)?StringBuilder'
    $sbMatches = [regex]::Matches($Content, $sbPattern)

    foreach ($sbMatch in $sbMatches) {
        $sbVarName = $sbMatch.Groups[1].Value
        $appendPattern = "(?i)$([regex]::Escape($sbVarName))\.(?:Append|AppendLine|AppendFormat)\s*\(\s*""([^""]*)""\s*\)"
        $appendMatches = [regex]::Matches($Content, $appendPattern)
        $fullSql = ""
        foreach ($appendMatch in $appendMatches) {
            $fullSql += " " + $appendMatch.Groups[1].Value
        }
        $fullSql = $fullSql.Trim()
        if ($fullSql -match '(?i)(SELECT|INSERT|UPDATE|DELETE|MERGE)') {
            [void]$sqlStrings.Add($fullSql)
        }
    }

    # パターン3b: With ブロック内の .Append（StringBuilder 変数名省略形）
    $withPattern = '(?is)With\s+(\w+)\s*\r?\n(.*?)End\s+With'
    $withMatches = [regex]::Matches($Content, $withPattern)
    foreach ($wm in $withMatches) {
        $body = $wm.Groups[2].Value
        $appendInWith = [regex]::Matches($body, '(?i)\.(?:Append|AppendLine|AppendFormat)\s*\(\s*"([^"]*)"\s*\)')
        $parts = @()
        foreach ($am in $appendInWith) {
            $parts += $am.Groups[1].Value
        }
        $fullSql = ($parts -join " ").Trim()
        if ($fullSql -match '(?i)(SELECT|INSERT|UPDATE|DELETE|MERGE)') {
            [void]$sqlStrings.Add($fullSql)
        }
    }

    # パターン4: VB.NET の行継続文字（ _）を使った複数行文字列
    $lineContinuation = '(?i)"([^"]*)"[\s]*_\s*\r?\n\s*(?:&\s*)?"([^"]*)"'
    $lcMatches = [regex]::Matches($Content, $lineContinuation)
    foreach ($match in $lcMatches) {
        $combined = $match.Groups[1].Value + " " + $match.Groups[2].Value
        if ($combined -match '(?i)(SELECT|INSERT|UPDATE|DELETE|MERGE)') {
            $exists = $false
            foreach ($existing in $sqlStrings) {
                if ($existing.Contains($match.Groups[1].Value)) {
                    $exists = $true
                    break
                }
            }
            if (-not $exists) {
                [void]$sqlStrings.Add($combined)
            }
        }
    }

    return $sqlStrings
}

function Get-VbNetClassAndMethods {
    param([string]$Content)

    $result = @{
        ClassName = ""
        Methods   = [System.Collections.ArrayList]::new()
    }

    if ($Content -match '(?i)(?:Public|Private|Friend|Protected)?\s*Class\s+(\w+)') {
        $result.ClassName = $Matches[1]
    }

    $methodPattern = '(?i)(?:Public|Private|Friend|Protected)?\s*(?:Shared\s+)?(?:Overrides\s+)?((?<!End\s)Function|(?<!End\s)Sub)\s+(\w+)'
    $methodMatches = [regex]::Matches($Content, $methodPattern)

    for ($i = 0; $i -lt $methodMatches.Count; $i++) {
        $startIdx = $methodMatches[$i].Index
        $endIdx = if ($i + 1 -lt $methodMatches.Count) { $methodMatches[$i + 1].Index } else { $Content.Length }
        $methodContent = $Content.Substring($startIdx, $endIdx - $startIdx)
        $methodName = $methodMatches[$i].Groups[2].Value

        [void]$result.Methods.Add(@{
            Name    = $methodName
            Content = $methodContent
        })
    }

    if ($result.Methods.Count -eq 0) {
        [void]$result.Methods.Add(@{
            Name    = "(FILE)"
            Content = $Content
        })
    }

    return $result
}

function Assert-SqlParserLoaded {
    if (-not (Get-Command Get-TableAndColumns -ErrorAction SilentlyContinue)) {
        throw "Parse-OracleSql.ps1 が先に読み込まれている必要があります（Get-TableAndColumns 関数が見つかりません）"
    }
}

function ConvertFrom-VbNetDacFile {
    param(
        [string]$FilePath,
        [System.Collections.Generic.HashSet[string]]$GlobalCteNames = $null,
        [string[]]$KnownCteNames = @()
    )

    Assert-SqlParserLoaded
    $content = [System.IO.File]::ReadAllText($FilePath, [System.Text.UTF8Encoding]::new($false))
    $fileName = [System.IO.Path]::GetFileName($FilePath)
    $classInfo = Get-VbNetClassAndMethods -Content $content

    $results = [System.Collections.ArrayList]::new()

    $fileCteNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $allLiterals = [regex]::Matches($content, '"([^"]*)"')
    $allLiteralText = ($allLiterals | ForEach-Object { $_.Groups[1].Value }) -join " "
    if ($allLiteralText -match '(?i)\bWITH\b') {
        $cteFound = [regex]::Matches($allLiteralText, '(?i)([\w$]+)\s+AS\s*\(')
        foreach ($cf in $cteFound) {
            [void]$fileCteNames.Add($cf.Groups[1].Value.ToUpper())
        }
    }
    if ($null -ne $GlobalCteNames) {
        foreach ($gCte in $GlobalCteNames) {
            [void]$fileCteNames.Add($gCte)
        }
    }
    foreach ($kn in $KnownCteNames) {
        if (-not [string]::IsNullOrWhiteSpace($kn)) {
            [void]$fileCteNames.Add($kn.Trim().ToUpper())
        }
    }

    foreach ($method in $classInfo.Methods) {
        $sqlStrings = Get-VbNetSqlStrings -Content $method.Content
        $featureName = "VB:$($classInfo.ClassName).$($method.Name)"

        foreach ($sql in $sqlStrings) {
            foreach ($opType in @("INSERT", "SELECT", "UPDATE", "DELETE", "MERGE")) {
                $extracted = Get-TableAndColumns -SqlFragment $sql -OperationType $opType -AdditionalCteNames @($fileCteNames)

                foreach ($item in $extracted) {
                    if ($fileCteNames.Count -gt 0 -and $fileCteNames.Contains($item.TableName)) { continue }
                    [void]$results.Add(@{
                        SourceType  = "VB.NET"
                        SourceFile  = $fileName
                        ObjectType  = "DAC"
                        ObjectName  = $classInfo.ClassName
                        ProcName    = $method.Name
                        FeatureName = $featureName
                        TableName   = $item.TableName
                        ColumnName  = $item.ColumnName
                        Operation   = $item.Operation
                    })
                }
            }
        }
    }

    return $results
}

function ConvertFrom-VbNetDacDirectory {
    param(
        [string]$SourcePath,
        [string]$DacFilePattern = "*dac*.vb",
        [string[]]$ExcludePatterns = @(),
        [string[]]$ExcludeTables = @(),
        [string[]]$ExcludeSchemas = @(),
        [string[]]$KnownCteNames = @()
    )

    Write-Host "[VB.NET] 解析開始: $SourcePath" -ForegroundColor Cyan

    $files = Get-ChildItem -Path $SourcePath -Filter $DacFilePattern -Recurse -File
    foreach ($pattern in $ExcludePatterns) {
        $files = $files | Where-Object { $_.Name -notlike $pattern }
    }

    Write-Host "[VB.NET] 対象DACファイル数: $($files.Count)" -ForegroundColor Cyan

    $globalCteNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($file in $files) {
        try {
            $fileContent = [System.IO.File]::ReadAllText($file.FullName, [System.Text.UTF8Encoding]::new($false))
            $literals = [regex]::Matches($fileContent, '"([^"]*)"')
            $literalText = ($literals | ForEach-Object { $_.Groups[1].Value }) -join " "
            if ($literalText -match '(?i)\bWITH\b') {
                $cteMatches = [regex]::Matches($literalText, '(?i)([\w$]+)\s+AS\s*\(')
                foreach ($cm in $cteMatches) {
                    [void]$globalCteNames.Add($cm.Groups[1].Value.ToUpper())
                }
            }
        } catch { }
    }
    if ($globalCteNames.Count -gt 0) {
        Write-Host "[VB.NET] WITH句(CTE)名を検出: $($globalCteNames.Count) 件 ($($globalCteNames -join ', '))" -ForegroundColor Cyan
    }

    $allResults = [System.Collections.ArrayList]::new()
    $fileCount = 0

    foreach ($file in $files) {
        $fileCount++
        Write-Progress -Activity "VB.NET DAC 解析中" -Status "$fileCount / $($files.Count): $($file.Name)" -PercentComplete (($fileCount / $files.Count) * 100)

        try {
            $fileResults = ConvertFrom-VbNetDacFile -FilePath $file.FullName -GlobalCteNames $globalCteNames -KnownCteNames $KnownCteNames

            foreach ($result in $fileResults) {
                $skip = $false
                if ($result.TableName -in $ExcludeTables) { $skip = $true }
                foreach ($schema in $ExcludeSchemas) {
                    if ($result.TableName -like "$schema*") { $skip = $true }
                }
                if (-not $skip) {
                    [void]$allResults.Add($result)
                }
            }
        }
        catch {
            Write-Warning "[VB.NET] 解析エラー: $($file.FullName) - $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "VB.NET DAC 解析中" -Completed
    Write-Host "[VB.NET] 解析完了: $($allResults.Count) 件のCRUDエントリを検出" -ForegroundColor Green

    return $allResults
}
