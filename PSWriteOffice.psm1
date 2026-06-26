# to speed up development adding direct path to binaries, instead of the the Lib folder
$DevelopmentBasePath = Join-Path (Join-Path (Join-Path $PSScriptRoot 'Sources') 'PSWriteOffice') 'bin'
$DevelopmentConfiguration = if ($env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION -in @('Debug', 'Release')) {
    $env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION
} elseif (Test-Path (Join-Path $DevelopmentBasePath 'Release')) {
    'Release'
} else {
    'Debug'
}
$DevelopmentPath = Join-Path $DevelopmentBasePath $DevelopmentConfiguration
$DevelopmentFolderCore = "net8.0"
$DevelopmentFolderDefault = "net472"
$DevelopmentFramework = if ($PSVersionTable.PSEdition -eq 'Core') {
    $DevelopmentFolderCore
} else {
    $DevelopmentFolderDefault
}
$DevelopmentBinaryPath = Join-Path (Join-Path $DevelopmentPath $DevelopmentFramework) 'PSWriteOffice.dll'
$Development = if ($env:PSWRITEOFFICE_USE_DEVELOPMENT_BINARIES -eq 'false') {
    $false
} else {
    Test-Path $DevelopmentBinaryPath
}
$BinaryModules = @(
    "PSWriteOffice.dll"
)
$AssemblyFolders = Get-ChildItem -Path (Join-Path $PSScriptRoot 'Lib') -Directory -ErrorAction SilentlyContinue

function Import-PSWriteOfficeDevelopmentBinaryModule {
    param(
        [Parameter(Mandatory)]
        [string] $Path
    )

    $loaderTypeName = 'PSWriteOffice.DevelopmentModuleLoadContext.ModuleAssemblyLoadContext'
    if (-not ($loaderTypeName -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;

namespace PSWriteOffice.DevelopmentModuleLoadContext;

public sealed class ModuleAssemblyLoadContext : AssemblyLoadContext
{
    private static readonly object Sync = new();
    private static readonly Dictionary<string, ModuleAssemblyLoadContext> Contexts = new(StringComparer.OrdinalIgnoreCase);
    private readonly string _assemblyDirectory;
    private readonly string _moduleAssemblyPath;
    private readonly AssemblyDependencyResolver _resolver;
    private Assembly _moduleAssembly;

    private ModuleAssemblyLoadContext(string moduleAssemblyPath, string contextName) : base(contextName, isCollectible: false)
    {
        _moduleAssemblyPath = Path.GetFullPath(moduleAssemblyPath);
        _assemblyDirectory = Path.GetDirectoryName(_moduleAssemblyPath) ?? string.Empty;
        _resolver = new AssemblyDependencyResolver(_moduleAssemblyPath);
    }

    public static Assembly LoadModule(string moduleAssemblyPath, string contextName)
    {
        if (string.IsNullOrWhiteSpace(moduleAssemblyPath))
        {
            throw new ArgumentException("Module assembly path is required.", nameof(moduleAssemblyPath));
        }

        string fullPath = Path.GetFullPath(moduleAssemblyPath);
        if (!File.Exists(fullPath))
        {
            throw new FileNotFoundException("Module assembly was not found.", fullPath);
        }

        lock (Sync)
        {
            if (!Contexts.TryGetValue(fullPath, out ModuleAssemblyLoadContext context))
            {
                context = new ModuleAssemblyLoadContext(fullPath, string.IsNullOrWhiteSpace(contextName) ? Path.GetFileNameWithoutExtension(fullPath) : contextName);
                Contexts[fullPath] = context;
            }

            return context.LoadMainModule();
        }
    }

    protected override Assembly Load(AssemblyName assemblyName)
    {
        if (assemblyName == null || string.IsNullOrWhiteSpace(assemblyName.Name))
        {
            return null;
        }

        AssemblyName loaderAssembly = typeof(ModuleAssemblyLoadContext).Assembly.GetName();
        if (AssemblyName.ReferenceMatchesDefinition(loaderAssembly, assemblyName))
        {
            return typeof(ModuleAssemblyLoadContext).Assembly;
        }

        if (string.Equals(assemblyName.Name, "System.Management.Automation", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        string assemblyPath = _resolver.ResolveAssemblyToPath(assemblyName);
        if (!string.IsNullOrWhiteSpace(assemblyPath) && File.Exists(assemblyPath))
        {
            return LoadFromAssemblyPath(assemblyPath);
        }

        string fallbackPath = Path.Combine(_assemblyDirectory, assemblyName.Name + ".dll");
        return File.Exists(fallbackPath) ? LoadFromAssemblyPath(fallbackPath) : null;
    }

    protected override IntPtr LoadUnmanagedDll(string unmanagedDllName)
    {
        string libraryPath = _resolver.ResolveUnmanagedDllToPath(unmanagedDllName);
        if (libraryPath != null)
        {
            return LoadUnmanagedDllFromPath(libraryPath);
        }

        return IntPtr.Zero;
    }

    private Assembly LoadMainModule()
    {
        if (_moduleAssembly == null)
        {
            _moduleAssembly = LoadFromAssemblyPath(_moduleAssemblyPath);
        }

        return _moduleAssembly;
    }
}
'@ -ErrorAction Stop
    }

    $importModule = Get-Command -Name Import-Module -Module Microsoft.PowerShell.Core
    $moduleAssembly = [PSWriteOffice.DevelopmentModuleLoadContext.ModuleAssemblyLoadContext]::LoadModule($Path, 'PSWriteOfficeDevelopment')
    $innerModule = & $importModule -Assembly $moduleAssembly -Force -PassThru -ErrorAction Stop

    if ($innerModule) {
        $addExportedCmdlet = [System.Management.Automation.PSModuleInfo].GetMethod(
            'AddExportedCmdlet',
            [System.Reflection.BindingFlags]'Instance, NonPublic'
        )
        if ($null -ne $addExportedCmdlet) {
            foreach ($cmdlet in $innerModule.ExportedCmdlets.Values) {
                $addExportedCmdlet.Invoke($ExecutionContext.SessionState.Module, @(, $cmdlet)) | Out-Null
            }

            $addExportedAlias = [System.Management.Automation.PSModuleInfo].GetMethod(
                'AddExportedAlias',
                [System.Reflection.BindingFlags]'Instance, NonPublic'
            )
            if ($null -ne $addExportedAlias) {
                foreach ($alias in $innerModule.ExportedAliases.Values) {
                    $aliasTarget = if ([string]::IsNullOrWhiteSpace($alias.Definition)) {
                        $alias.ResolvedCommandName
                    } else {
                        $alias.Definition
                    }

                    Set-Alias -Name $alias.Name -Value $aliasTarget -Scope Local -Force -ErrorAction Stop
                    $exportedAlias = $ExecutionContext.SessionState.InvokeCommand.GetCommand($alias.Name, [System.Management.Automation.CommandTypes]::Alias)
                    if ($null -ne $exportedAlias) {
                        $addExportedAlias.Invoke($ExecutionContext.SessionState.Module, @(, $exportedAlias)) | Out-Null
                    }
                }
            }
        } else {
            throw 'AddExportedCmdlet is not available on this PowerShell version.'
        }
    }
}

# ensure script file collections always exist (legacy folders were removed)
if (-not (Test-Path variable:Classes)) { $Classes = @() }
if (-not (Test-Path variable:Enums)) { $Enums = @() }
if (-not (Test-Path variable:Private)) { $Private = @() }
if (-not (Test-Path variable:Public)) { $Public = @() }

# Lets find which libraries we need to load
$Default = $false
$Core = $false
$Standard = $false
foreach ($A in $AssemblyFolders.Name) {
    if ($A -eq 'Default') {
        $Default = $true
    } elseif ($A -eq 'Core') {
        $Core = $true
    } elseif ($A -eq 'Standard') {
        $Standard = $true
    }
}
if ($Standard -and $Core -and $Default) {
    $FrameworkNet = 'Default'
    $Framework = 'Standard'
} elseif ($Standard -and $Core) {
    $Framework = 'Standard'
    $FrameworkNet = 'Standard'
} elseif ($Core -and $Default) {
    $Framework = 'Core'
    $FrameworkNet = 'Default'
} elseif ($Standard -and $Default) {
    $Framework = 'Standard'
    $FrameworkNet = 'Default'
} elseif ($Standard) {
    $Framework = 'Standard'
    $FrameworkNet = 'Standard'
} elseif ($Core) {
    $Framework = 'Core'
    $FrameworkNet = ''
} elseif ($Default) {
    $Framework = ''
    $FrameworkNet = 'Default'
} else {
    #Write-Error -Message 'No assemblies found'
}

# Ensure native runtime libraries are discoverable on Windows
if ($IsWindows) {
    $arch = [System.Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture
    $archFolder = switch ($arch) {
        'X64' {
            'win-x64'
        }
        'X86' {
            'win-x86'
        }
        'Arm64' {
            'win-arm64'
        }
        'Arm' {
            'win-arm'
        }
        default {
            'win-x64'
        }
    }

    if ($Development) {
        $baseDir = if ($PSEdition -eq 'Core') {
            Join-Path $DevelopmentPath $DevelopmentFolderCore
        } else {
            Join-Path $DevelopmentPath $DevelopmentFolderDefault
        }
    } else {
        $baseDir = if ($PSEdition -eq 'Core') {
            Join-Path $PSScriptRoot "Lib/$Framework"
        } elseif ($FrameworkNet) {
            Join-Path $PSScriptRoot "Lib/$FrameworkNet"
        } else {
            $null
        }
    }

    if ($baseDir) {
        $runtimePath = Join-Path $baseDir "runtimes/$archFolder/native"
        if (Test-Path $runtimePath) {
            Write-Verbose -Message "Adding $runtimePath to PATH"
            $env:PATH = "$runtimePath;" + $env:PATH
        }
    }
}

$BinaryDev = if ($Development) {
    @(
        foreach ($BinaryModule in $BinaryModules) {
            if ($PSEdition -eq 'Core') {
                $Variable = Resolve-Path (Join-Path (Join-Path $DevelopmentPath $DevelopmentFolderCore) $BinaryModule)
            } else {
                $Variable = Resolve-Path (Join-Path (Join-Path $DevelopmentPath $DevelopmentFolderDefault) $BinaryModule)
            }
            $Variable
            Write-Verbose "Development mode: Using binaries from $Variable"
        }
    )
} else {
    @()
}

$ImportedBinaryModules = @()
$FoundErrors = @(
    if ($Development) {
        foreach ($BinaryModule in $BinaryDev) {
            try {
                $binaryModulePath = (Resolve-Path -LiteralPath $BinaryModule).ProviderPath
                if ($PSEdition -eq 'Core') {
                    Import-PSWriteOfficeDevelopmentBinaryModule -Path $binaryModulePath
                } else {
                    Import-Module -Name $BinaryModule -Force -ErrorAction Stop
                }
            } catch {
                Write-Warning "Failed to import module $($BinaryModule): $($_.Exception.Message)"
                $true
            }
        }
    } else {
        foreach ($BinaryModule in $BinaryModules) {
            try {
                if ($Framework -and $PSEdition -eq 'Core') {
                    $importedModule = Import-Module -Name "$PSScriptRoot\Lib\$Framework\$BinaryModule" -Force -PassThru -ErrorAction Stop
                    if ($importedModule) {
                        $ImportedBinaryModules += $importedModule
                    }
                }
                if ($FrameworkNet -and $PSEdition -ne 'Core') {
                    $importedModule = Import-Module -Name "$PSScriptRoot\Lib\$FrameworkNet\$BinaryModule" -Force -PassThru -ErrorAction Stop
                    if ($importedModule) {
                        $ImportedBinaryModules += $importedModule
                    }
                }
            } catch {
                Write-Warning "Failed to import module $($BinaryModule): $($_.Exception.Message)"
                $true
            }
        }
    }
    #Dot source the files
    foreach ($Import in @($Classes + $Enums + $Private + $Public)) {
        try {
            . $Import.Fullname
        } catch {
            Write-Error -Message "Failed to import functions from $($import.Fullname): $_"
            $true
        }
    }
)

if ($FoundErrors.Count -gt 0) {
    $ModuleName = (Get-ChildItem $PSScriptRoot\*.psd1).BaseName
    Write-Warning "Importing module $ModuleName failed. Fix errors before continuing."
    throw "Importing module $ModuleName failed. Fix errors before continuing."
    #break
}

$cmdletsForAliasExport = @(
    $ExecutionContext.SessionState.Module.ExportedCmdlets.Values
    foreach ($importedModule in $ImportedBinaryModules) {
        $importedModule.ExportedCmdlets.Values
    }
) | Where-Object {
    $null -ne $_ -and $null -ne $_.ImplementingType
} | Sort-Object -Property Name -Unique

foreach ($cmdlet in $cmdletsForAliasExport) {
    $aliasAttributes = $cmdlet.ImplementingType.GetCustomAttributes([System.Management.Automation.AliasAttribute], $true)
    foreach ($attribute in $aliasAttributes) {
        foreach ($aliasName in $attribute.AliasNames) {
            if ([string]::IsNullOrWhiteSpace($aliasName)) {
                continue
            }

            Set-Alias -Name $aliasName -Value $cmdlet.Name -Scope Local -Force -ErrorAction Stop
        }
    }
}

Export-ModuleMember -Alias '*' -Cmdlet '*'
