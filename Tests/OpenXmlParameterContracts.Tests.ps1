BeforeAll {
    $ModuleManifest = if ($env:PSWRITEOFFICE_MODULE_MANIFEST) {
        $env:PSWRITEOFFICE_MODULE_MANIFEST
    } else {
        Join-Path $PSScriptRoot '..\PSWriteOffice.psd1'
    }
    Import-Module $ModuleManifest -Global -ErrorAction Stop

    function Get-ParameterLeafType {
        param([Parameter(Mandatory)][Type] $Type)

        $nullableType = [Nullable]::GetUnderlyingType($Type)
        if ($null -ne $nullableType) {
            Get-ParameterLeafType -Type $nullableType
            return
        }

        if ($Type.HasElementType) {
            Get-ParameterLeafType -Type ($Type.GetElementType())
            return
        }

        if ($Type.IsGenericType) {
            foreach ($argumentType in $Type.GetGenericArguments()) {
                Get-ParameterLeafType -Type $argumentType
            }
            return
        }

        $Type
    }
}

Describe 'Open XML parameter contracts' {
    It 'does not expose Open XML SDK value structs as cmdlet parameters' {
        $leaks = Get-Command -Module PSWriteOffice | ForEach-Object {
            $command = $_
            $_.Parameters.GetEnumerator() | ForEach-Object {
                $parameter = $_
                Get-ParameterLeafType -Type $parameter.Value.ParameterType |
                    Where-Object { $_.FullName -like 'DocumentFormat.OpenXml*' } |
                    ForEach-Object { "$($command.Name):$($parameter.Key):$($_.FullName)" }
            }
        }

        @($leaks) | Should -BeNullOrEmpty
    }

    It 'uses OfficeIMO-owned CLR enums for PowerShell-facing document values' {
        $contracts = @{
            'Add-OfficeWordSection:BreakType' = 'OfficeIMO.Word.WordSectionBreakType'
            'Add-OfficeWordBreak:BreakType' = 'OfficeIMO.Word.WordBreakType'
            'Add-OfficeWordHeader:Type' = 'OfficeIMO.Word.WordHeaderFooterType'
            'Add-OfficeWordFooter:Type' = 'OfficeIMO.Word.WordHeaderFooterType'
            'Add-OfficeWordParagraph:Alignment' = 'OfficeIMO.Word.WordParagraphAlignment'
            'Add-OfficeWordText:Underline' = 'OfficeIMO.Word.WordUnderlineStyle'
            'New-OfficeWordTableCell:UnderlineStyle' = 'OfficeIMO.Word.WordUnderlineStyle'
            'New-OfficeWordTableCell:Align' = 'OfficeIMO.Word.WordParagraphAlignment'
            'New-OfficeWordTableCell:VerticalAlign' = 'OfficeIMO.Word.WordTableVerticalAlignment'
            'Add-OfficeWordTextBox:HorizontalPositionRelativeFrom' = 'OfficeIMO.Word.WordHorizontalRelativePosition'
            'Add-OfficeWordTextBox:VerticalPositionRelativeFrom' = 'OfficeIMO.Word.WordVerticalRelativePosition'
            'Protect-OfficeWordDocument:ProtectionType' = 'OfficeIMO.Word.WordDocumentProtectionType'
            'Set-OfficeWordTableCell:TextDirection' = 'OfficeIMO.Word.WordTextDirection'
            'Add-OfficePowerPointSlide:LayoutType' = 'OfficeIMO.PowerPoint.PowerPointSlideLayoutType'
            'Set-OfficePowerPointSlideLayout:LayoutType' = 'OfficeIMO.PowerPoint.PowerPointSlideLayoutType'
        }

        foreach ($entry in $contracts.GetEnumerator()) {
            $commandName, $parameterName = $entry.Key -split ':', 2
            $parameterType = (Get-Command $commandName).Parameters[$parameterName].ParameterType
            $nullableType = [Nullable]::GetUnderlyingType($parameterType)
            if ($null -ne $nullableType) {
                $parameterType = $nullableType
            }
            $parameterType.FullName | Should -Be $entry.Value
            $parameterType.IsEnum | Should -BeTrue
        }
    }
}
