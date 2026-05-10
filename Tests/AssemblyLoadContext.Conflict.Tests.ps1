Describe 'Packaged AssemblyLoadContext conflict isolation' {
    It 'loads after a default-context Open XML assembly without sharing the default ALC' {
        $packagedModuleRoot = Join-Path $PSScriptRoot '..\Artefacts\Unpacked\Modules'
        $packagedModule = Join-Path $packagedModuleRoot 'PSWriteOffice'
        $packagedLoader = Join-Path $packagedModule 'Lib\Core\PSWriteOffice.ModuleLoadContext.dll'
        $conflictOpenXmlPath = $env:PSWRITEOFFICE_CONFLICT_OPENXML_PATH

        if ($PSVersionTable.PSEdition -ne 'Core' -or
            -not (Test-Path -LiteralPath $packagedLoader) -or
            [string]::IsNullOrWhiteSpace($conflictOpenXmlPath) -or
            -not (Test-Path -LiteralPath $conflictOpenXmlPath)) {
            Set-ItResult -Skipped -Because 'packaged Core artifact and conflict Open XML assembly are required'
            return
        }

        $moduleRootLiteral = $packagedModuleRoot.Replace("'", "''")
        $conflictOpenXmlLiteral = $conflictOpenXmlPath.Replace("'", "''")
        $script = @"
`$ErrorActionPreference = 'Stop'
`$WarningPreference = 'SilentlyContinue'
`$moduleRoot = '$moduleRootLiteral'
`$conflictOpenXmlPath = '$conflictOpenXmlLiteral'
`$env:PSModulePath = `$moduleRoot + [IO.Path]::PathSeparator + `$env:PSModulePath

Add-Type -Path `$conflictOpenXmlPath -ErrorAction Stop
`$defaultOpenXmlAssembly = [AppDomain]::CurrentDomain.GetAssemblies() |
    Where-Object { `$_.GetName().Name -eq 'DocumentFormat.OpenXml' } |
    Select-Object -First 1
`$defaultOpenXmlAlc = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext(`$defaultOpenXmlAssembly)

Import-Module PSWriteOffice -Force
`$command = Get-Command New-OfficeWord -ErrorAction Stop
`$commandAssembly = `$command.ImplementingType.Assembly
`$commandAlc = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext(`$commandAssembly)
`$loadedAssemblies = [System.Runtime.Loader.AssemblyLoadContext]::All |
    ForEach-Object {
        `$alc = `$_
        foreach (`$assembly in `$alc.Assemblies) {
            if (`$assembly.GetName().Name -in @('PSWriteOffice', 'OfficeIMO.Word', 'DocumentFormat.OpenXml')) {
                [pscustomobject]@{
                    Assembly = `$assembly.GetName().Name
                    Version = `$assembly.GetName().Version.ToString()
                    ALC = `$alc.Name
                    IsDefault = [object]::ReferenceEquals(`$alc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
                    Location = `$assembly.Location
                }
            }
        }
    }

[pscustomobject]@{
    DefaultOpenXmlAssembly = `$defaultOpenXmlAssembly.Location
    DefaultOpenXmlVersion = `$defaultOpenXmlAssembly.GetName().Version.ToString()
    DefaultOpenXmlALC = `$defaultOpenXmlAlc.Name
    DefaultOpenXmlALCIsDefault = [object]::ReferenceEquals(`$defaultOpenXmlAlc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
    NewOfficeWordAssembly = `$commandAssembly.Location
    NewOfficeWordALC = `$commandAlc.Name
    NewOfficeWordALCIsDefault = [object]::ReferenceEquals(`$commandAlc, [System.Runtime.Loader.AssemblyLoadContext]::Default)
    LoadedAssemblies = @(`$loadedAssemblies)
} | ConvertTo-Json -Depth 6 -Compress
"@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
        $output = pwsh -NoProfile -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)

        $json = $output | Where-Object { $_ -is [string] -and $_.TrimStart().StartsWith('{') } | Select-Object -Last 1
        $json | Should -Not -BeNullOrEmpty -Because ($output -join [Environment]::NewLine)
        $result = $json | ConvertFrom-Json

        $result.DefaultOpenXmlAssembly | Should -Be $conflictOpenXmlPath
        $result.DefaultOpenXmlALCIsDefault | Should -BeTrue
        $result.NewOfficeWordAssembly | Should -BeLike '*\Artefacts\Unpacked\Modules\PSWriteOffice\Lib\Core\PSWriteOffice.dll'
        $result.NewOfficeWordALC | Should -Be 'PSWriteOffice'
        $result.NewOfficeWordALCIsDefault | Should -BeFalse

        $loadedAssemblies = @($result.LoadedAssemblies)
        $psWriteOfficeAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'PSWriteOffice' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $officeImoAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'OfficeIMO.Word' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $moduleOpenXmlAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'DocumentFormat.OpenXml' -and $_.ALC -eq 'PSWriteOffice' } | Select-Object -First 1
        $defaultOpenXmlAssembly = $loadedAssemblies | Where-Object { $_.Assembly -eq 'DocumentFormat.OpenXml' -and $_.IsDefault } | Select-Object -First 1

        $psWriteOfficeAssembly.IsDefault | Should -BeFalse
        $officeImoAssembly.IsDefault | Should -BeFalse
        $moduleOpenXmlAssembly.IsDefault | Should -BeFalse
        $defaultOpenXmlAssembly.IsDefault | Should -BeTrue
    }
}
