function New-OfficeWordText {
    [cmdletBinding()]
    param(
        [OfficeIMO.Word.WordDocument] $Document,
        [OfficeIMO.Word.WordParagraph] $Paragraph,
        [string[]]$Text,
        [nullable[bool][]] $Bold,
        [nullable[bool][]] $Italic,
        [nullable[DocumentFormat.OpenXml.Wordprocessing.UnderlineValues][]] $Underline,
        [string[]] $Color,
        [DocumentFormat.OpenXml.Wordprocessing.JustificationValues] $Alignment,
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
            $Paragraph.Underline = $Underline[$T]
        }
        if ($Color -and $Color.Count -ge $T -and $Color[$T]) {
            $ColorToSet = (ConvertFrom-Color -Color $Color[$T])
            if ($ColorToSet) {
                $Paragraph.Color = $ColorToSet
            }
        }
        if ($Alignment) {
            $Paragraph.ParagraphAlignment = $Alignment
        }
    }
    if ($ReturnObject) {
        $Paragraph
    }
}

Register-ArgumentCompleter -CommandName New-OfficeWordText -ParameterName Color -ScriptBlock $Script:ScriptBlockColors