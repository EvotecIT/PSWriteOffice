function New-OfficeWordText {
    [cmdletBinding(DefaultParameterSetName = 'Document')]
    param(
        [Parameter(ParameterSetName = 'Document')][OfficeIMO.Word.WordDocument] $Document,
        [Parameter(ParameterSetName = 'Paragraph')][OfficeIMO.Word.WordParagraph] $Paragraph,

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
            $Paragraph.Underline = $Underline[$T]
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
            $Paragraph.ParagraphAlignment = $Alignment
        }
    }
    if ($ReturnObject) {
        $Paragraph
    }
}

Register-ArgumentCompleter -CommandName New-OfficeWordText -ParameterName Color -ScriptBlock $Script:ScriptBlockColors