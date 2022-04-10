function New-OfficeWordListItem {
    [cmdletBinding()]
    param(
        [OfficeIMO.Word.WordList] $List,
        [int] $Level,
        [string[]]$Text,
        [nullable[bool][]] $Bold,
        [nullable[bool][]] $Italic,
        [nullable[DocumentFormat.OpenXml.Wordprocessing.UnderlineValues][]] $Underline,
        [string[]] $Color,
        [nullable[DocumentFormat.OpenXml.Wordprocessing.JustificationValues]] $Alignment,
        [switch] $Suppress
    )
    if ($List) {
        # This is standard usage + internal function
        $ListItem = $List.AddItem($Text, $Level)
        if (-not $Suppress) {
            $ListItem
        }
    } else {
        # This is to be used when use within New-OfficeWordList
        [ordered] @{
            List      = $null
            Level     = $Level
            Text      = $Text
            Bold      = $Bold
            Italic    = $Italic
            Underline = $Underline
            Color     = $Color
            Alignment = $Alignment
            Suppress  = $Suppress
        }
    }
}