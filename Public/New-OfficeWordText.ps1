function New-OfficeWordText {
    [cmdletBinding()]
    param(
        $Document,
        [string[]]$Text,
        [DocumentFormat.OpenXml.SpaceProcessingModeValues] $Space = [DocumentFormat.OpenXml.SpaceProcessingModeValues]::Preserve,
        [nullable[bool][]] $Bold,
        [nullable[bool][]] $Italic,
        [nullable[bool][]] $Underline,
        [string[]] $Color,
        [DocumentFormat.OpenXml.Wordprocessing.JustificationValues] $Alignment
    )
    for ($T = 0; $T -le $Text.Count; $T++) {
        $WordText = [DocumentFormat.OpenXml.Wordprocessing.Text] @{
            Text  = $Text[$T]
            Space = $Space
        }
        if ($Space -and $Space.Count -ge $T -and $Space[$T]) {
            $WordText.Space = $Space[$T]
        }

        $Run = [DocumentFormat.OpenXml.Wordprocessing.Run]::new()
        $RunProperties = [DocumentFormat.OpenXml.Wordprocessing.RunProperties]::new()

        # Setting up bold
        if ($Bold -and $Bold.Count -ge $T -and $Bold[$T]) {
            $RunProperties.Bold = [DocumentFormat.OpenXml.Wordprocessing.Bold]::new()
        }
        if ($Italic -and $Italic.Count -ge $T -and $Italic[$T]) {
            $RunProperties.Italic = [DocumentFormat.OpenXml.Wordprocessing.Italic]::new()
        }
        if ($Underline -and $Underline.Count -ge $T -and $Underline[$T]) {
            $RunProperties.Underline = [DocumentFormat.OpenXml.Wordprocessing.Underline]::new()
        }
        if ($Color -and $Color.Count -ge $T -and $Color[$T]) {
            $ColorToSet = (ConvertFrom-Color -Color $Color[$T])
            if ($ColorToSet) {
                $RunProperties.Color = [DocumentFormat.OpenXml.Wordprocessing.Color]::new()
                $RunProperties.Color.Val = $ColorToSet
            }
        }

        if ($Alignment) {
            # Alignement applies to whole paragraph so we only assign it once
            $ParagraphProperties = [DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties]::new()
            $ParagraphProperties.Justification = [DocumentFormat.OpenXml.Wordprocessing.Justification] @{
                Val = $Alignment
            }
        }

        $Run.AppendChild($RunProperties)
        $Run.AppendChild($WordText)
        if (-not $Paragraph) {
            $Paragraph = [DocumentFormat.OpenXml.Wordprocessing.Paragraph]::new()
            if ($ParagraphProperties) {
                # Paragraph properties apply only on first run
                $Paragraph.Append($ParagraphProperties)
            }
            $Paragraph.Append($Run)
            $Paragraph = $Document.MainDocumentPart.Document.Body.AppendChild($Paragraph)
        } else {
            $Paragraph = $Paragraph.AppendChild($Run)
        }
    }
}