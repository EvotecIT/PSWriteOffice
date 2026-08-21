BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    $script:OfficeCommands = @(Get-Command -Module PSWriteOffice -CommandType Cmdlet)
}

Describe 'PSWriteOffice public API consistency' {
    It 'uses Path as the canonical source-file parameter name' {
        $filePathCommands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('FilePath') })
        $inputPathCommands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('InputPath') })

        $filePathCommands.Name | Should -BeNullOrEmpty
        $inputPathCommands.Name | Should -Be @('Export-OfficeDocumentPdf')
    }

    It 'reserves Show for semantic visibility choices' {
        $allowed = @(
            'Set-OfficeExcelGridlines'
            'Set-OfficeExcelSheetVisibility'
        )
        $commands = @($script:OfficeCommands | Where-Object { $_.Parameters.ContainsKey('Show') })

        @($commands.Name | Sort-Object) | Should -Be @($allowed | Sort-Object)
    }

    It 'keeps format conversion out of New and Save lifecycle parameters' {
        $lifecycle = @($script:OfficeCommands | Where-Object { $_.Verb -in 'New', 'Save' })

        @($lifecycle | Where-Object { $_.Parameters.ContainsKey('PdfPath') }).Name | Should -BeNullOrEmpty
        @($lifecycle | Where-Object { $_.Verb -eq 'New' -and $_.Parameters.ContainsKey('AutoSave') }).Name | Should -BeNullOrEmpty
    }

    It 'offers PassThru on commands that save artifacts' {
        $missing = @($script:OfficeCommands | Where-Object { $_.Verb -eq 'Save' -and -not $_.Parameters.ContainsKey('PassThru') })

        $missing.Name | Should -BeNullOrEmpty
    }

    It 'offers PassThru on commands that perform mutations' {
        $mutationVerbs = @('Add', 'Clear', 'Copy', 'Edit', 'Move', 'Protect', 'Remove', 'Rename', 'Set', 'Unprotect', 'Update')
        $intentionalTransforms = @('Set-OfficeConfluenceManagedSection')
        $missing = @(
            $script:OfficeCommands |
                Where-Object {
                    $_.Verb -in $mutationVerbs -and
                    $_.Name -notin $intentionalTransforms -and
                    -not $_.Parameters.ContainsKey('PassThru')
                }
        )

        $missing.Name | Should -BeNullOrEmpty
    }

    It 'exposes one format-neutral PDF export command' {
        $command = Get-Command Export-OfficeDocumentPdf

        $command.Parameters.Keys | Should -Contain 'InputPath'
        $command.Parameters.Keys | Should -Contain 'Path'
        $command.Parameters.Keys | Should -Contain 'Document'
        $command.Parameters.Keys | Should -Contain 'PassThru'
        $command.Parameters.Keys | Should -Contain 'Open'
        $command.ParameterSets.Name | Should -Contain 'Document'
        $command.ParameterSets.Name | Should -Contain 'Path'
    }

    It 'uses Path and DestinationPath for file copies' {
        $command = Get-Command Copy-OfficeExcelWorkbook

        $command.Parameters.Keys | Should -Contain 'Path'
        $command.Parameters.Keys | Should -Contain 'DestinationPath'
        $command.Parameters.Keys | Should -Not -Contain 'InputPath'
        $command.Parameters['Path'].Aliases | Should -Contain 'InputPath'
        $command.Parameters['DestinationPath'].Aliases | Should -Contain 'OutputPath'
    }

    It 'keeps saved New commands quiet unless PassThru is requested' {
        $cases = @(
            @{ Name = 'Word'; Extension = 'docx' }
            @{ Name = 'Excel'; Extension = 'xlsx' }
            @{ Name = 'PowerPoint'; Extension = 'pptx' }
            @{ Name = 'Markdown'; Extension = 'md' }
            @{ Name = 'PDF'; Extension = 'pdf' }
            @{ Name = 'Visio'; Extension = 'vsdx' }
        )

        foreach ($case in $cases) {
            $quietPath = Join-Path $TestDrive ("quiet-{0}.{1}" -f $case.Name, $case.Extension)
            $quietOutput = @(switch ($case.Name) {
                    Word { New-OfficeWord -Path $quietPath }
                    Excel { New-OfficeExcel -Path $quietPath }
                    PowerPoint { New-OfficePowerPoint -Path $quietPath }
                    Markdown { New-OfficeMarkdown -Path $quietPath }
                    PDF { New-OfficePdf -Path $quietPath { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $quietPath }
                })
            $quietOutput | Should -HaveCount 0 -Because "$($case.Name) should be quiet by default"
            Test-Path -LiteralPath $quietPath | Should -BeTrue

            $passThruPath = Join-Path $TestDrive ("passthru-{0}.{1}" -f $case.Name, $case.Extension)
            $passThruOutput = @(switch ($case.Name) {
                    Word { New-OfficeWord -Path $passThruPath -PassThru }
                    Excel { New-OfficeExcel -Path $passThruPath -PassThru }
                    PowerPoint { New-OfficePowerPoint -Path $passThruPath -PassThru }
                    Markdown { New-OfficeMarkdown -Path $passThruPath -PassThru }
                    PDF { New-OfficePdf -Path $passThruPath -PassThru { PdfParagraph 'API contract' } }
                    Visio { New-OfficeVisio -Path $passThruPath -PassThru }
                })
            $passThruOutput | Should -HaveCount 1 -Because "$($case.Name) PassThru should emit the saved file"
            $passThruOutput[0] | Should -BeOfType System.IO.FileInfo
            $passThruOutput[0].FullName | Should -Be $passThruPath
        }
    }

    It 'does not teach suppression of PSWriteOffice output in public examples' {
        $publicRoots = @(
            Join-Path $PSScriptRoot '..\Examples'
            Join-Path $PSScriptRoot '..\Website\content'
        )
        $offenders = [System.Collections.Generic.List[string]]::new()

        $publicFiles = @(
            foreach ($root in $publicRoots) {
                Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.md' }
            }
        )
        $publicFiles | ForEach-Object {
            $file = $_
            $snippets = if ($file.Extension -eq '.ps1') {
                @([pscustomobject]@{ Text = [System.IO.File]::ReadAllText($file.FullName); LineOffset = 0 })
            } else {
                $markdown = [System.IO.File]::ReadAllText($file.FullName)
                @(
                    [regex]::Matches($markdown, '(?ms)^```powershell\s*\r?\n(?<code>.*?)^```\s*$') | ForEach-Object {
                        $prefix = $markdown.Substring(0, $_.Groups['code'].Index)
                        [pscustomobject]@{
                            Text = $_.Groups['code'].Value
                            LineOffset = @($prefix -split '\r?\n').Count - 1
                        }
                    }
                )
            }

            foreach ($snippet in $snippets) {
                $tokens = $null
                $errors = $null
                $ast = [System.Management.Automation.Language.Parser]::ParseInput($snippet.Text, [ref] $tokens, [ref] $errors)
                if ($file.Extension -eq '.ps1') {
                    $errors | Should -BeNullOrEmpty
                }

                foreach ($pipeline in $ast.FindAll({
                        param($node)
                        $node -is [System.Management.Automation.Language.PipelineAst]
                    }, $true)) {
                    if ($pipeline.PipelineElements.Count -lt 2) {
                        continue
                    }

                    $last = $pipeline.PipelineElements[-1]
                    if ($last -isnot [System.Management.Automation.Language.CommandAst] -or $last.GetCommandName() -ne 'Out-Null') {
                        continue
                    }

                    $commandNames = @(
                        $pipeline.PipelineElements |
                            Where-Object { $_ -is [System.Management.Automation.Language.CommandAst] } |
                            ForEach-Object { $_.GetCommandName() }
                    )
                    if ($commandNames | Where-Object { $_ -match 'Office' }) {
                        $line = $snippet.LineOffset + $pipeline.Extent.StartLineNumber
                        $offenders.Add(('{0}:{1}' -f $file.FullName, $line))
                    }
                }
            }
        }

        $offenders | Should -BeNullOrEmpty
    }
}
