function New-OfficeWordText {
    [cmdletBinding(DefaultParameterSetName = 'Document')]
    param(
        [Parameter(ParameterSetName = 'Paragraph')]
        [Parameter(ParameterSetName = 'Document', Mandatory)][OfficeIMO.Word.WordDocument] $Document,

        [Parameter(ParameterSetName = 'Paragraph', Mandatory)]
        [OfficeIMO.Word.WordParagraph] $Paragraph,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [string[]]$Text,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [nullable[bool][]] $Bold,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [nullable[bool][]] $Italic,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [nullable[DocumentFormat.OpenXml.Wordprocessing.UnderlineValues][]] $Underline,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [string[]] $Color,

        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [DocumentFormat.OpenXml.Wordprocessing.JustificationValues] $Alignment,
        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [OfficeIMO.Word.WordParagraphStyles] $Style,
        [Parameter(ParameterSetName = 'Document')]
        [Parameter(ParameterSetName = 'Paragraph')]
        [switch] $ReturnObject
    )
    if (-not $Paragraph) {
        $Paragraph = $Document.AddParagraph()
    }
    for ($T = 0; $T -lt $Text.Count; $T++) {
        $Paragraph = $Paragraph.AddText($Text[$T])

        if ($Bold -and $Bold.Count -ge $T -and $Bold[$T]) {
            $Paragraph.Bold = $Bold[$T]
        }
        if ($Italic -and $Italic.Count -ge $T -and $Italic[$T]) {
            $Paragraph.Italic = $Italic[$T]
        }
        if ($Underline -and $Underline.Count -ge $T -and $Underline[$T]) {

            #Write-Host "Value of Single: $([DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::Single)"
            #Write-Host "Type of Single: $([DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::Single.GetType().FullName)"

            # Write-Host "Value of Underline[`$T]: $($Underline[$T])"
            # Write-Host "Type of Underline[`$T]: $($Underline[$T].GetType().FullName)"

            #Write-Color -Text "This works"
            #$Paragraph.Underline = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::Single # This works
            #$Test = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues] "single"
            #$Paragraph.Underline = $Test # this works
            #$Test2 = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues] "$($Underline[$T].Value.ToLower())"
            #$Paragraph.Underline = $Test2 # this works

            $Paragraph.Underline = $Underline[$T].Value.ToLower()

            #$Paragraph.Underline = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::new($Underline[$T].Value.Value)
            #$Paragraph.Underline = $Underline[$T] # This doesn't, even tho same type
            #Write-Color -Text "This doesn't"
            #$underlineValue = $Underline[$T]
            # if ($underlineValue -ne $null) {
            #     # Ensure the value is of the correct type
            #     $underlineValue = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::new($underlineValue.Value)
            # }
            #if ($underlineValue -ne $null) {
            # Ensure the value is of the correct type
            #$underlineValue = [DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]::new([DocumentFormat.OpenXml.Wordprocessing.UnderlineValues]$underlineValue.Value)
            # }

            # $Paragraph.Underline = $underlineValue


            #$Paragraph.Underline = $Underline[$T]
            #$Paragraph.SetUnderline($Underline[$T])

            #$Paragraph.Underline = $Underline[$T].Value
        }
        if ($Color -and $Color.Count -ge $T -and $Color[$T]) {
            $ColorToSet = (ConvertFrom-Color -Color $Color[$T])
            if ($ColorToSet) {
                $Paragraph.Color = $ColorToSet
            }
        }
        if ($Style) {
            $Paragraph.Style = $Style
        }
        if ($Alignment) {
            $Paragraph.ParagraphAlignment = $Alignment.Value.ToLower()
        }
    }
    if ($ReturnObject) {
        $Paragraph
    }
}

Register-ArgumentCompleter -CommandName New-OfficeWordText -ParameterName Color -ScriptBlock $Script:ScriptBlockColors