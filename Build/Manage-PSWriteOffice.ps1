Clear-Host
Import-Module 'C:\Support\GitHub\PSPublishModule\PSPublishModule.psd1' -Force


Invoke-ModuleBuild -ModuleName 'PSWriteOffice' {
    # Usual defaults as per standard module
    $Manifest = [ordered] @{
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

        DotNetFrameworkVersion = '4.7.2'
    }
    New-ConfigurationManifest @Manifest

    New-ConfigurationModule -Type ExternalModule -Name 'Microsoft.PowerShell.Utility', 'Microsoft.PowerShell.Management'
    New-ConfigurationModule -Type RequiredModule -Name 'PSSharedGoods' -Version Latest -Guid Auto
    New-ConfigurationModule -Type ApprovedModule -Name 'PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword', 'PSPublishModule'
    New-ConfigurationModuleSkip -IgnoreFunctionName 'Select-Unique'

    $ConfigurationFormat = [ordered] @{
        RemoveComments                              = $false

        PlaceOpenBraceEnable                        = $true
        PlaceOpenBraceOnSameLine                    = $true
        PlaceOpenBraceNewLineAfter                  = $true
        PlaceOpenBraceIgnoreOneLineBlock            = $false

        PlaceCloseBraceEnable                       = $true
        PlaceCloseBraceNewLineAfter                 = $true
        PlaceCloseBraceIgnoreOneLineBlock           = $false
        PlaceCloseBraceNoEmptyLineBefore            = $true

        UseConsistentIndentationEnable              = $true
        UseConsistentIndentationKind                = 'space'
        UseConsistentIndentationPipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
        UseConsistentIndentationIndentationSize     = 4

        UseConsistentWhitespaceEnable               = $true
        UseConsistentWhitespaceCheckInnerBrace      = $true
        UseConsistentWhitespaceCheckOpenBrace       = $true
        UseConsistentWhitespaceCheckOpenParen       = $true
        UseConsistentWhitespaceCheckOperator        = $true
        UseConsistentWhitespaceCheckPipe            = $true
        UseConsistentWhitespaceCheckSeparator       = $true

        AlignAssignmentStatementEnable              = $true
        AlignAssignmentStatementCheckHashtable      = $true

        UseCorrectCasingEnable                      = $true
    }
    # format PSD1 and PSM1 files when merging into a single file
    # enable formatting is not required as Configuration is provided
    New-ConfigurationFormat -ApplyTo 'OnMergePSM1', 'OnMergePSD1' -Sort None @ConfigurationFormat
    # format PSD1 and PSM1 files within the module
    # enable formatting is required to make sure that formatting is applied (with default settings)
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'DefaultPSM1' -EnableFormatting -Sort None
    # when creating PSD1 use special style without comments and with only required parameters
    New-ConfigurationFormat -ApplyTo 'DefaultPSD1', 'OnMergePSD1' -PSD1Style 'Minimal'
    # configuration for documentation, at the same time it enables documentation processing
    New-ConfigurationDocumentation -Enable:$false -StartClean -UpdateWhenNew -PathReadme 'Docs\Readme.md' -Path 'Docs'

    New-ConfigurationImportModule -ImportSelf

    New-ConfigurationBuild -Enable:$true -SignModule -MergeModuleOnBuild -MergeFunctionsFromApprovedModules -CertificateThumbprint '36A8A2D0E227D81A2D3B60DCE0CFCF23BEFC343B' -ResolveBinaryConflicts -ResolveBinaryConflictsName 'PSWriteOffice' -NETProjectName 'PSWriteOffice' -NETConfiguration Release -NETFramework 'netstandard2.0', 'net472'

    New-ConfigurationArtefact -Type Unpacked -Enable -Path "$PSScriptRoot\..\Artefacts\Unpacked" -ModulesPath "$PSScriptRoot\..\Artefacts\Unpacked\Modules" -RequiredModulesPath "$PSScriptRoot\..\Artefacts\Unpacked\Modules" -AddRequiredModules
    New-ConfigurationArtefact -Type Packed -Enable -Path "$PSScriptRoot\..\Artefacts\Packed" -ArtefactName '<ModuleName>.v<ModuleVersion>.zip'

    # global options for publishing to github/psgallery
    #New-ConfigurationPublish -Type PowerShellGallery -FilePath 'C:\Support\Important\PowerShellGalleryAPI.txt' -Enabled:$true
    #New-ConfigurationPublish -Type GitHub -FilePath 'C:\Support\Important\GitHubAPI.txt' -UserName 'EvotecIT' -Enabled:$true
}

# $Configuration = @{
#     Information = @{
#         ModuleName       = 'PSWriteOffice'

#         LibrariesCore    = 'Lib\Core'
#         LibrariesDefault = 'Lib\Default'

#         Manifest         = @{
#             # Minimum version of the Windows PowerShell engine required by this module
#             PowerShellVersion      = '5.1'
#             # prevent using over CORE/PS 7
#             CompatiblePSEditions   = @('Desktop', 'Core')
#             # ID used to uniquely identify this module
#             GUID                   = 'd75a279d-30c2-4c2d-ae0d-12f1f3bf4d39'
#             # Version number of this module.
#             ModuleVersion          = '0.0.X'
#             # Author of this module
#             Author                 = 'Przemyslaw Klys'
#             # Company or vendor of this module
#             CompanyName            = 'Evotec'
#             # Copyright statement for this module
#             Copyright              = "(c) 2011 - $((Get-Date).Year) Przemyslaw Klys @ Evotec. All rights reserved."
#             # Description of the functionality provided by this module
#             Description            = 'Experimental PowerShell Module to create and edit Microsoft Word, Microsoft Excel, and Microsoft PowerPoint documents without having Microsoft Office installed.'
#             # Tags applied to this module. These help with module discovery in online galleries.
#             Tags                   = @('word', 'docx', 'write', 'PSWord', 'office', 'windows', 'doc', 'pswriteword', 'linux', 'macos')
#             # A URL to the main website for this project.
#             ProjectUri             = 'https://github.com/EvotecIT/PSWriteOffice'

#             IconUri                = 'https://evotec.xyz/wp-content/uploads/2018/10/PSWriteWord.png'

#             LicenseUri             = 'https://github.com/EvotecIT/PSWriteOffice/blob/master/License'

#             RequiredModules        = @(
#                 @{ ModuleName = 'PSSharedGoods'; ModuleVersion = "Latest"; Guid = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe' }
#             )
#             DotNetFrameworkVersion = '4.7.2'
#         }
#     }
#     Options     = @{
#         Merge             = @{
#             Sort           = 'None'
#             FormatCodePSM1 = @{
#                 Enabled           = $true
#                 RemoveComments    = $false
#                 FormatterSettings = @{
#                     IncludeRules = @(
#                         'PSPlaceOpenBrace',
#                         'PSPlaceCloseBrace',
#                         'PSUseConsistentWhitespace',
#                         'PSUseConsistentIndentation',
#                         'PSAlignAssignmentStatement',
#                         'PSUseCorrectCasing'
#                     )

#                     Rules        = @{
#                         PSPlaceOpenBrace           = @{
#                             Enable             = $true
#                             OnSameLine         = $true
#                             NewLineAfter       = $true
#                             IgnoreOneLineBlock = $true
#                         }

#                         PSPlaceCloseBrace          = @{
#                             Enable             = $true
#                             NewLineAfter       = $false
#                             IgnoreOneLineBlock = $true
#                             NoEmptyLineBefore  = $false
#                         }

#                         PSUseConsistentIndentation = @{
#                             Enable              = $true
#                             Kind                = 'space'
#                             PipelineIndentation = 'IncreaseIndentationAfterEveryPipeline'
#                             IndentationSize     = 4
#                         }

#                         PSUseConsistentWhitespace  = @{
#                             Enable          = $true
#                             CheckInnerBrace = $true
#                             CheckOpenBrace  = $true
#                             CheckOpenParen  = $true
#                             CheckOperator   = $true
#                             CheckPipe       = $true
#                             CheckSeparator  = $true
#                         }

#                         PSAlignAssignmentStatement = @{
#                             Enable         = $true
#                             CheckHashtable = $true
#                         }

#                         PSUseCorrectCasing         = @{
#                             Enable = $true
#                         }
#                     }
#                 }
#             }
#             FormatCodePSD1 = @{
#                 Enabled        = $true
#                 RemoveComments = $false
#             }
#             Integrate      = @{
#                 ApprovedModules = @('PSSharedGoods', 'PSWriteColor', 'Connectimo', 'PSUnifi', 'PSWebToolbox', 'PSMyPassword')
#             }
#         }
#         Standard          = @{
#             FormatCodePSM1 = @{

#             }
#             FormatCodePSD1 = @{
#                 Enabled = $true
#                 #RemoveComments = $true
#             }
#         }
#         ImportModules     = @{
#             Self            = $true
#             RequiredModules = $false
#             Verbose         = $false
#         }
#         PowerShellGallery = @{
#             ApiKey   = 'C:\Support\Important\PowerShellGalleryAPI.txt'
#             FromFile = $true
#         }
#         GitHub            = @{
#             ApiKey   = 'C:\Support\Important\GithubAPI.txt'
#             FromFile = $true
#             UserName = 'EvotecIT'
#             #RepositoryName = 'PSWriteHTML'
#         }
#         Documentation     = @{
#             Path       = 'Docs'
#             PathReadme = 'Docs\Readme.md'
#         }
#     }
#     Steps       = @{
#         BuildLibraries     = @{
#             Enable        = $true # build once every time nuget gets updated
#             Configuration = 'Release'
#             Framework     = 'netstandard2.0', 'net472'
#             ProjectName   = 'PSWriteOffice'
#         }
#         BuildModule        = @{  # requires Enable to be on to process all of that
#             Enable                 = $true
#             DeleteBefore           = $true
#             Merge                  = $true
#             MergeMissing           = $true
#             LibrarySeparateFile    = $false
#             LibraryDotSource       = $true
#             ClassesDotSource       = $false
#             SignMerged             = $true
#             CreateFileCatalog      = $false # not working
#             Releases               = $true
#             ReleasesUnpacked       = $false
#             RefreshPSD1Only        = $false
#             ResolveBinaryConflicts = @{
#                 ProjectName = 'PSWriteOffice'
#             }
#             DebugDLL               = $true
#         }
#         BuildDocumentation = $true
#         ImportModules      = @{
#             Self            = $true
#             RequiredModules = $false
#             Verbose         = $false
#         }
#         PublishModule      = @{  # requires Enable to be on to process all of that
#             Enabled      = $false
#             Prerelease   = ''
#             RequireForce = $false
#             GitHub       = $false
#         }
#     }
# }

# New-PrepareModule -Configuration $Configuration