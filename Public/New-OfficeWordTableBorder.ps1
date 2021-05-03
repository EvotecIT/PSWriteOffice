function New-OfficeWordTableBorder {
    [cmdletBinding()]
    param(

    )

    # Table borders
    $TableBorders = [DocumentFormat.OpenXml.Wordprocessing.TableBorders]::new()
    $TopBorder = [DocumentFormat.OpenXml.Wordprocessing.TopBorder] @{
        Val  = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size = [DocumentFormat.OpenXml.UInt32Value]::new(24)
    }
    $BottomBorder = [DocumentFormat.OpenXml.Wordprocessing.BottomBorder] @{
        Val  = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size = [DocumentFormat.OpenXml.UInt32Value]::new(24)
    }
    $LeftBorder = [DocumentFormat.OpenXml.Wordprocessing.LeftBorder] @{
        Val  = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size = [DocumentFormat.OpenXml.UInt32Value]::new(24)
    }
    $RightBorder = [DocumentFormat.OpenXml.Wordprocessing.RightBorder] @{
        Val  = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size = [DocumentFormat.OpenXml.UInt32Value]::new(24)
    }
    $InsideHorizontalBorder = [DocumentFormat.OpenXml.Wordprocessing.InsideHorizontalBorder] @{
        Val  = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size = [DocumentFormat.OpenXml.UInt32Value]::new(24)
    }
    $InsideVerticalBorder = [DocumentFormat.OpenXml.Wordprocessing.InsideVerticalBorder] @{
        Val   = [DocumentFormat.OpenXml.Wordprocessing.BorderValues]::Dashed
        Size  = [DocumentFormat.OpenXml.UInt32Value]::new(24)
        #Space = [DocumentFormat.OpenXml.UInt32Value] "0U"
        Color = "red"
        #Frame #DocumentFormat.OpenXml.OnOffValue Frame {get;set;}
        # Shadow # DocumentFormat.OpenXml.OnOffValue Shadow {get;set;}
        # DocumentFormat.OpenXml.EnumValue[DocumentFormat.OpenXml.Wordprocessing.ThemeColorValues] ThemeColor {get;set;}
        # DocumentFormat.OpenXml.StringValue ThemeShade {get;set;}
        # DocumentFormat.OpenXml.StringValue ThemeTint {get;set;}
    }

    $TableBorders.BottomBorder = $BottomBorder
    $TableBorders.LeftBorder = $LeftBorder
    $TableBorders.RightBorder = $RightBorder
    $TableBorders.TopBorder = $TopBorder
    $TableBorders.InsideHorizontalBorder = $InsideHorizontalBorder
    $TableBorders.InsideVerticalBorder = $InsideVerticalBorder
    , $TableBorders
}