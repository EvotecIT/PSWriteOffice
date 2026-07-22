BeforeAll {
    $script:repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
    $script:catalogScript = Join-Path $script:repoRoot 'Build\Export-WebsiteDocumentationCatalog.ps1'
    $script:catalogPath = Join-Path $script:repoRoot 'WebsiteArtifacts\documentation\command-catalog.json'
    $script:projectManifestPath = Join-Path $script:repoRoot 'WebsiteArtifacts\project-manifest.json'
    $script:moduleManifestPath = Join-Path $script:repoRoot 'PSWriteOffice.psd1'
    $script:apiRoot = Join-Path $script:repoRoot 'WebsiteArtifacts\apidocs\powershell'
}

Describe 'PSWriteOffice website documentation catalog' {
    It 'covers every exported cmdlet exactly once' {
        $outputPath = Join-Path $TestDrive 'command-catalog.json'
        & $script:catalogScript -RepositoryRoot $script:repoRoot -OutputPath $outputPath | Out-Null

        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $catalog = Get-Content -LiteralPath $outputPath -Raw | ConvertFrom-Json

        $catalog.module.commandCount | Should -Be @($module.CmdletsToExport).Count
        $catalog.module.aliasCount | Should -Be @($module.AliasesToExport).Count
        $catalog.module.familyCount | Should -Be @($catalog.families).Count
        (@($catalog.families | Measure-Object -Property commandCount -Sum).Sum) | Should -Be $catalog.module.commandCount
        @($catalog.families | Where-Object commandCount -LT 1).Count | Should -Be 0
    }

    It 'keeps the committed catalog deterministic and current' {
        $outputPath = Join-Path $TestDrive 'command-catalog.json'
        & $script:catalogScript -RepositoryRoot $script:repoRoot -OutputPath $outputPath | Out-Null

        (Get-Content -LiteralPath $outputPath -Raw).Trim() |
            Should -Be (Get-Content -LiteralPath $script:catalogPath -Raw).Trim()
    }

    It 'publishes real docs, examples, and API surfaces at the module version' {
        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $project = Get-Content -LiteralPath $script:projectManifestPath -Raw | ConvertFrom-Json

        $project.version | Should -Be ([string] $module.ModuleVersion)
        $project.surfaces.docs | Should -BeTrue
        $project.surfaces.examples | Should -BeTrue
        $project.surfaces.apiPowerShell | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.docs) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.examples) | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:repoRoot $project.artifacts.documentationCatalog) | Should -BeTrue
    }

    It 'keeps the committed PowerShell API bundle aligned with every exported cmdlet' {
        $module = Import-PowerShellDataFile -LiteralPath $script:moduleManifestPath
        $expected = @($module.CmdletsToExport) | Sort-Object -Unique
        $apiManifest = Import-PowerShellDataFile -LiteralPath (Join-Path $script:apiRoot 'PSWriteOffice.psd1')
        [xml] $help = Get-Content -LiteralPath (Join-Path $script:apiRoot 'PSWriteOffice-help.xml') -Raw
        $metadata = Get-Content -LiteralPath (Join-Path $script:apiRoot 'command-metadata.json') -Raw | ConvertFrom-Json

        $manifestCommands = @($apiManifest.CmdletsToExport) | Sort-Object -Unique
        $helpCommands = @($help.helpItems.command | ForEach-Object { [string] $_.details.name }) | Sort-Object -Unique
        $metadataCommands = @($metadata.commands | ForEach-Object { [string] $_.name }) | Sort-Object -Unique

        $manifestCommands | Should -Be $expected
        $helpCommands | Should -Be $expected
        $metadataCommands | Should -Be $expected
        @($metadata.commands | Where-Object { -not $_.sourcePath }).Count | Should -Be 0
    }
}
