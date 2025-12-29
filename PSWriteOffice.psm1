# to speed up development adding direct path to binaries, instead of the the Lib folder
$Development = $true
$DevelopmentPath = "$PSScriptRoot\Sources\PSWriteOffice\bin\Debug"
$DevelopmentFolderCore = "net8.0"
$DevelopmentFolderDefault = "net472"
$BinaryModules = @(
    "PSWriteOffice.dll"
)

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
            Write-Warning -Message "Adding $runtimePath to PATH"
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

$Assembly = @(
    if ($Development) {
        if ($PSEdition -eq 'Core') {
            Get-ChildItem -Path $DevelopmentPath\$DevelopmentFolderCore -Filter '*.dll' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]runtimes[\\/]' -and $_.Name -ne 'Microsoft.Bcl.AsyncInterfaces.dll' }
        } else {
            Get-ChildItem -Path $DevelopmentPath\$DevelopmentFolderDefault -Filter '*.dll' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]runtimes[\\/]' -and $_.Name -ne 'Microsoft.Bcl.AsyncInterfaces.dll' }
        }
    } else {
        if ($Framework -and $PSEdition -eq 'Core') {
            Get-ChildItem -Path $PSScriptRoot\Lib\$Framework -Filter '*.dll' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]runtimes[\\/]' -and $_.Name -ne 'Microsoft.Bcl.AsyncInterfaces.dll' }
        }
        if ($FrameworkNet -and $PSEdition -ne 'Core') {
            Get-ChildItem -Path $PSScriptRoot\Lib\$FrameworkNet -Filter '*.dll' -Recurse | Where-Object { $_.FullName -notmatch '[\\/]runtime(s[\\/]' -and $_.Name -ne 'Microsoft.Bcl.AsyncInterfaces.dll' }
        }
    }
)

$BinaryDev = @(
    foreach ($BinaryModule in $BinaryModules) {
        if ($PSEdition -eq 'Core') {
            $Variable = Resolve-Path "$DevelopmentPath\$DevelopmentFolderCore\$BinaryModule"
        } else {
            $Variable = Resolve-Path "$DevelopmentPath\$DevelopmentFolderDefault\$BinaryModule"
        }
        $Variable
        Write-Warning "Development mode: Using binaries from $Variable"
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
    foreach ($Import in @($Assembly)) {
        try {
            Write-Verbose -Message $Import.FullName
            Add-Type -Path $Import.Fullname -ErrorAction Stop
            #  }
        } catch [System.Reflection.ReflectionTypeLoadException] {
            Write-Warning "Processing $($Import.Name) Exception: $($_.Exception.Message)"
            $LoaderExceptions = $($_.Exception.LoaderExceptions) | Sort-Object -Unique
            foreach ($E in $LoaderExceptions) {
                Write-Warning "Processing $($Import.Name) LoaderExceptions: $($E.Message)"
            }
            $true
            #Write-Error -Message "StackTrace: $($_.Exception.StackTrace)"
        } catch {
            Write-Warning "Processing $($Import.Name) Exception: $($_.Exception.Message)"
            $LoaderExceptions = $($_.Exception.LoaderExceptions) | Sort-Object -Unique
            foreach ($E in $LoaderExceptions) {
                Write-Warning "Processing $($Import.Name) LoaderExceptions: $($E.Message)"
            }
            $true
            #Write-Error -Message "StackTrace: $($_.Exception.StackTrace)"
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
