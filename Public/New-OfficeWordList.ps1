function New-OfficeWordList {
    [cmdletBinding()]
    param(
        [ScriptBlock] $Content,
        [OfficeImo.Word.WordDocument] $Document,
        [OfficeIMO.Word.WordListStyle] $Style = [OfficeIMO.Word.WordListStyle]::Bulleted,
        [switch] $Suppress
    )

    $List = $Document.AddList($Style)
    if ($Content) {
        $ListItems = & $Content
        foreach ($Item in $ListItems) {
            # We will use the same function we use externally but internall
            # But we define the list now
            $Item.List = $List
            # We also don't want to have output from List Items
            $Item.Suppress = $true
            New-OfficeWordListItem @Item
        }
    }
    if (-not $Suppress) {
        $List
    }
}