function New-OfficeWordTableLook {
    [cmdletBinding()]
    param(


    )
    #TableLook tableLook = new TableLook() { Val = "04A0" };
    $TableLook = [DocumentFormat.OpenXml.Wordprocessing.TableLook] @{
        Val              = "04A0"
        FirstRow         = $true
        LastRow          = $false
        FirstColumn      = $true
        LastColumn       = $false
        NoHorizontalBand = $false
        NoVerticalBand   = $true
    }
    , $TableLook

}