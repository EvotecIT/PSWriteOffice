BeforeDiscovery {
    $repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
    $examplesRoot = Join-Path $repoRoot 'Examples'
    $exampleCases = Get-ChildItem -LiteralPath $examplesRoot -Recurse -File -Filter '*.ps1' |
        Sort-Object FullName |
        ForEach-Object {
            $relativePath = $_.FullName.Substring($repoRoot.Length)
            $relativePath = $relativePath.TrimStart([char[]] @(
                [System.IO.Path]::DirectorySeparatorChar,
                [System.IO.Path]::AltDirectorySeparatorChar
            ))

            @{
                Path = $_.FullName
                RelativePath = $relativePath
            }
        }

    $recipeCases = $exampleCases | Where-Object { [System.IO.Path]::GetFileName($_.Path) -like 'Recipe-*' }
}

Describe 'Repository example scripts' {
    It 'parses <RelativePath>' -ForEach $exampleCases {
        $tokens = $null
        $parseErrors = $null
        [System.Management.Automation.Language.Parser]::ParseFile(
            $Path,
            [ref] $tokens,
            [ref] $parseErrors
        ) | Out-Null

        @($parseErrors | ForEach-Object Message) | Should -BeNullOrEmpty
    }

    It 'keeps recipe <RelativePath> focused on the document workflow' -ForEach $recipeCases {
        $content = Get-Content -LiteralPath $Path -Raw

        $content | Should -Not -Match '\A\s*param\s*\('
        $content | Should -Not -Match '(?m)^\s*(\$ErrorActionPreference\s*=|Import-Module\b|New-Item\b)'
        $content | Should -Not -Match '\b(Out-Null|Write-Host|Format-List)\b'
        $content | Should -Not -Match '\[Array\]::CreateInstance|\.Dispose\(\)'
    }
}
