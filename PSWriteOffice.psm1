# to speed up development adding direct path to binaries, instead of the the Lib folder
$DevelopmentPath = Join-Path $PSScriptRoot 'Sources\PSWriteOffice\bin\Debug'
$Development = Test-Path $DevelopmentPath
$DevelopmentFolderCore = "net8.0"
$DevelopmentFolderDefault = "net472"
$BinaryModules = @(
    "PSWriteOffice.dll"
)
$AssemblyFolders = Get-ChildItem -Path (Join-Path $PSScriptRoot 'Lib') -Directory -ErrorAction SilentlyContinue

# ensure script file collections always exist (legacy folders were removed)
if (-not (Test-Path variable:Classes)) { $Classes = @() }
if (-not (Test-Path variable:Enums)) { $Enums = @() }
if (-not (Test-Path variable:Private)) { $Private = @() }
if (-not (Test-Path variable:Public)) { $Public = @() }

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

$BinaryDev = @(
    foreach ($BinaryModule in $BinaryModules) {
        if ($PSEdition -eq 'Core') {
            $Variable = Resolve-Path "$DevelopmentPath\$DevelopmentFolderCore\$BinaryModule"
        } else {
            $Variable = Resolve-Path "$DevelopmentPath\$DevelopmentFolderDefault\$BinaryModule"
        }
        $Variable
        Write-Verbose "Development mode: Using binaries from $Variable"
    }
)

$FoundErrors = @(
    if ($Development) {
        foreach ($BinaryModule in $BinaryDev) {
            try {
                Import-Module -Name $BinaryModule -Force -ErrorAction Stop
            } catch {
                Write-Warning "Failed to import module $($BinaryModule): $($_.Exception.Message)"
                $true
            }
        }
    } else {
        foreach ($BinaryModule in $BinaryModules) {
            try {
                if ($Framework -and $PSEdition -eq 'Core') {
                    Import-Module -Name "$PSScriptRoot\Lib\$Framework\$BinaryModule" -Force -ErrorAction Stop
                }
                if ($FrameworkNet -and $PSEdition -ne 'Core') {
                    Import-Module -Name "$PSScriptRoot\Lib\$FrameworkNet\$BinaryModule" -Force -ErrorAction Stop
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

Export-ModuleMember -Function '*' -Alias '*' -Cmdlet '*'
