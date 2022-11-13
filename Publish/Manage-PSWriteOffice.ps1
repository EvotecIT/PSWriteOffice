﻿Clear-Host
Import-Module 'C:\Support\GitHub\PSPublishModule\PSPublishModule.psd1' -Force

$Configuration = @{
    Information = @{
        ModuleName           = 'PSWriteOffice'

        DirectoryProjects    = 'C:\Support\GitHub'
        DirectoryModulesCore = "$Env:USERPROFILE\Documents\PowerShell\Modules"
        DirectoryModules     = "$Env:USERPROFILE\Documents\WindowsPowerShell\Modules"

        FunctionsToExport    = 'Public'
        AliasesToExport      = 'Public'

        LibrariesCore        = 'Lib\Core'
        LibrariesDefault     = 'Lib\Default'

        Manifest             = @{
            # Minimum version of the Windows PowerShell engine required by this module
            PowerShellVersion      = '5.1'
            # prevent using over CORE/PS 7
            CompatiblePSEditions   = @('Desktop', 'Core')
            # ID used to uniquely identify this module
            GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
            # Version number of this module.
            ModuleVersion          = '0.0.X'
            # Author of this module
            Author                 = 'Przemyslaw Klys'
            # Company or vendor of this module
            CompanyName            = 'Evotec'
            # Copyright statement for this module
            Copyright              = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
            # Description of the functionality provided by this module
            Description            = 'Experimental PowerShell Module to create and edit Microsoft Word, Microsoft Excel, and Microsoft PowerPoint documents without having Microsoft Office installed.'
            # Tags applied to this module. These help with module discovery in online galleries.
            Tags                   = @('word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc', 'pswriteword', 'linux', 'macos')
            # A URL to the main website for this project.
            ProjectUri             = 'https://github.com/EvotecIT/PSWriteOffice'

            IconUri                = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteWord.png'

            LicenseUri             = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'

            RequiredModules        = @(
                @{ ModuleName = 'PSSharedGoods'; ModuleVersion = "Latest"; Guid = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe' }
            )
            DotNetFrameworkVersion = '4.7.2'
        }
    }
    Options     = @{
        Merge             = @{
            Sort           = 'None'
            FormatCodePSM1 = @{
                Enabled           = $true
                RemoveComments    = $true
                FormatterSettings = @{
                    IncludeRules = @(
                        'PSPlaceOpenBrace',
                        'PSPlaceCloseBrace',
                        'PSUseConsistentWhitespace',
                        'PSUseConsistentIndentation',
                        'PSAlignAssignmentStatement',
                        'PSUseCorrectCasing'
                    )

                    Rules        = @{
                        PSPlaceOpenBrace           = @{
                            Enable             = $true
                            OnSameLine         = $true
                            NewLineAfter       = $true
                            IgnoreOneLineBlock = $true
                        }

                        PSPlaceCloseBrace          = @{
                            Enable             = $true
                            NewLineAfter       = $false
                            IgnoreOneLineBlock = $true
                            NoEmptyLineBefore  = $false
                        }

                        PSUseConsistentIndentation = @{
                            Enable              = $true
                            Kind                = 'space'
                            PipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
                            IndentationSize     = 4
                        }

                        PSUseConsistentWhitespace  = @{
                            Enable          = $true
                            CheckInnerBrace = $true
                            CheckOpenBrace  = $true
                            CheckOpenParen  = $true
                            CheckOperator   = $true
                            CheckPipe       = $true
                            CheckSeparator  = $true
                        }

                        PSAlignAssignmentStatement = @{
                            Enable         = $true
                            CheckHashtable = $true
                        }

                        PSUseCorrectCasing         = @{
                            Enable = $true
                        }
                    }
                }
            }
            FormatCodePSD1 = @{
                Enabled        = $true
                RemoveComments = $false
            }
            Integrate      = @{
                ApprovedModules = @('PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword')
            }
        }
        Standard          = @{
            FormatCodePSM1 = @{

            }
            FormatCodePSD1 = @{
                Enabled = $true
                #RemoveComments = $true
            }
        }
        ImportModules     = @{
            Self            = $true
            RequiredModules = $false
            Verbose         = $false
        }
        PowerShellGallery = @{
            ApiKey   = 'C:\Support\Important\PowerShellGalleryAPI.txt'
            FromFile = $true
        }
        GitHub            = @{
            ApiKey   = 'C:\Support\Important\GithubAPI.txt'
            FromFile = $true
            UserName = 'EvotecIT'
            #RepositoryName = 'PSWriteHTML'
        }
        Documentation     = @{
            Path       = 'Docs'
            PathReadme = 'Docs\Readme.md'
        }
    }
    Steps       = @{
        <#
        BuildModule        = @{  # requires Enable to be on to process all of that
            Enable              = $true
            DeleteBefore        = $true
            Merge               = $true
            LibrarySeparateFile = $false
            MergeMissing        = $true
            Releases            = $true
            ReleasesUnpacked    = $false
            RefreshPSD1Only     = $false
        }
        #>
        BuildModule        = @{  # requires Enable to be on to process all of that
            Enable              = $true
            DeleteBefore        = $true
            Merge               = $true
            MergeMissing        = $true
            LibrarySeparateFile = $true
            LibraryDotSource    = $false
            ClassesDotSource    = $false
            SignMerged          = $true
            CreateFileCatalog   = $false # not working
            Releases            = $true
            ReleasesUnpacked    = $false
            RefreshPSD1Only     = $false
        }
        BuildDocumentation = $true
        ImportModules      = @{
            Self            = $true
            RequiredModules = $false
            Verbose         = $false
        }
        PublishModule      = @{  # requires Enable to be on to process all of that
            Enabled      = $false
            Prerelease   = ''
            RequireForce = $false
            GitHub       = $false
        }
    }
}

New-PrepareModule -Configuration $Configuration