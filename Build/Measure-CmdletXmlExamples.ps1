param(
    [string] $CmdletRoot = "$PSScriptRoot\..\Sources\PSWriteOffice\Cmdlets",
    [string] $DocsRoot = "$PSScriptRoot\..\Docs",
    [switch] $Summary,
    [switch] $RequireExamples
)

$ErrorActionPreference = 'Stop'

function ConvertTo-CommandName {
    param(
        [Parameter(Mandatory)]
        [string] $ClassBaseName
    )

    $officeIndex = $ClassBaseName.IndexOf('Office', [StringComparison]::Ordinal)
    if ($officeIndex -le 0) {
        return $ClassBaseName
    }

    return $ClassBaseName.Substring(0, $officeIndex) + '-' + $ClassBaseName.Substring($officeIndex)
}

function Get-XmlExampleMetric {
    param(
        [Parameter(Mandatory)]
        [string] $SourcePath,
        [Parameter(Mandatory)]
        [string] $RepositoryRoot,
        [Parameter(Mandatory)]
        [string] $DocsRoot
    )

    $lines = Get-Content -LiteralPath $SourcePath
    $cmdletLineIndexes = @(
        for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
            if ($lines[$lineIndex] -match '^\s*\[Cmdlet\(') {
                $lineIndex
            }
        }
    )

    if ($cmdletLineIndexes.Count -eq 0) {
        return
    }

    $category = Split-Path (Split-Path $SourcePath -Parent) -Leaf

    foreach ($cmdletLineIndex in $cmdletLineIndexes) {
        $className = $null
        $classSearchEnd = [Math]::Min($cmdletLineIndex + 12, $lines.Count - 1)
        for ($lineIndex = $cmdletLineIndex + 1; $lineIndex -le $classSearchEnd; $lineIndex++) {
            if ($lines[$lineIndex] -match 'public\s+(?:sealed\s+)?class\s+([A-Za-z0-9_]+)Command\b') {
                $className = $Matches[1]
                break
            }
        }

        if (-not $className) {
            continue
        }

        $commentEndIndex = $cmdletLineIndex - 1
        $commentStartIndex = $commentEndIndex
        while ($commentStartIndex -ge 0 -and $lines[$commentStartIndex] -match '^\s*///') {
            $commentStartIndex--
        }

        $comments = if ($commentStartIndex -lt $commentEndIndex) {
            $lines[($commentStartIndex + 1)..$commentEndIndex] -join "`n"
        } else {
            ''
        }

        $commandName = ConvertTo-CommandName -ClassBaseName $className
        $docPath = Join-Path $DocsRoot ($commandName + '.md')
        $examples = [regex]::Matches($comments, '(?s)///\s*<example>(.*?)///\s*</example>')
        $codeLineCounts = [System.Collections.Generic.List[int]]::new()
        $hasContextualExample = $false

        foreach ($example in $examples) {
            $body = $example.Groups[1].Value
            $codeMatches = [regex]::Matches($body, '(?s)<code>(.*?)</code>')

            foreach ($codeMatch in $codeMatches) {
                $code = [System.Net.WebUtility]::HtmlDecode($codeMatch.Groups[1].Value).Trim()
                $lineCount = @($code -split "`r?`n" | Where-Object { $_.Trim().Length -gt 0 }).Count
                $codeLineCounts.Add($lineCount)

                if ($lineCount -ge 2 -or $code -match '\$\w+\s*=|\{|\|') {
                    $hasContextualExample = $true
                }
            }
        }

        [pscustomobject]@{
            Category               = $category
            Command                = $commandName
            Source                 = [System.IO.Path]::GetRelativePath($RepositoryRoot, $SourcePath)
            HasDoc                 = Test-Path -LiteralPath $docPath -PathType Leaf
            ExampleCount           = $examples.Count
            MultilineExampleCount  = @($codeLineCounts | Where-Object { $_ -ge 3 }).Count
            MaxCodeLines           = if ($codeLineCounts.Count -gt 0) { ($codeLineCounts | Measure-Object -Maximum).Maximum } else { 0 }
            HasContextualExample   = $hasContextualExample
            NeedsXmlExampleWork    = $examples.Count -eq 0 -or -not $hasContextualExample
        }
    }
}

$resolvedCmdletRoot = (Resolve-Path -LiteralPath $CmdletRoot).Path
$repositoryRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$resolvedDocsRoot = (Resolve-Path -LiteralPath $DocsRoot).Path

$metrics = @(
    Get-ChildItem -LiteralPath $resolvedCmdletRoot -Recurse -Filter *.cs |
        ForEach-Object {
            Get-XmlExampleMetric -SourcePath $_.FullName -RepositoryRoot $repositoryRoot -DocsRoot $resolvedDocsRoot
        } |
        Sort-Object Category, Command
)

if ($Summary.IsPresent) {
    $metrics |
        Group-Object Category |
        Sort-Object Name |
        ForEach-Object {
            $group = $_.Group
            [pscustomobject]@{
                Category              = $_.Name
                Cmdlets               = $group.Count
                MissingDocs           = @($group | Where-Object { -not $_.HasDoc }).Count
                NoXmlExamples         = @($group | Where-Object { $_.ExampleCount -eq 0 }).Count
                NeedsXmlExampleWork   = @($group | Where-Object NeedsXmlExampleWork).Count
                WithMultilineExamples = @($group | Where-Object { $_.MultilineExampleCount -gt 0 }).Count
            }
        }
} elseif (-not $RequireExamples.IsPresent) {
    $metrics
}

if ($RequireExamples.IsPresent) {
    $missing = @($metrics | Where-Object NeedsXmlExampleWork)
    if ($missing.Count -gt 0) {
        throw "$($missing.Count) cmdlets still need contextual XML examples."
    }
}
