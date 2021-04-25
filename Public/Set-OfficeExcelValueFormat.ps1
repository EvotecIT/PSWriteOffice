function Set-OfficeExcelValueStyle {
    [cmdletBinding()]
    param(
        $Worksheet,
        [int] $Row,
        [int] $Column,
        [string] $Format,
        [int] $FormatID,
        #[string] $Color,
        #[string] $BackGroundColor,
        [nullable[bool]] $Bold, #         Property   bool Bold {get;set;}
        [ClosedXML.Excel.XLFontCharSet] $FontCharSet, # Property   ClosedXML.Excel.XLFontCharSet FontCharSet {get;set;}
        [alias('Color')][string] $FontColor, # Property   ClosedXML.Excel.XLColor FontColor {get;set;}
        [string] $BackGroundColor,
        [ClosedXML.Excel.XLFillPatternValues] $PatternType,

        [ClosedXML.Excel.XLFontFamilyNumberingValues] $FontFamilyNumbering, # Property   ClosedXML.Excel.XLFontFamilyNumberingValues FontFamilyNumbering {get;set;}
        [string] $FontName, # Property   string FontName {get;set;}
        [double] $FontSize, # Property   double FontSize {get;set;}
        [nullable[bool]] $Italic , # Property   bool Italic {get;set;}
        [nullable[bool]] $Shadow, # Property   bool Shadow {get;set;}
        [nullable[bool]] $Strikethrough , # Property   bool Strikethrough {get;set;}
        [ClosedXML.Excel.XLFontUnderlineValues] $Underline, # Property   ClosedXML.Excel.XLFontUnderlineValues Underline {get;set;}
        [ClosedXML.Excel.XLFontVerticalTextAlignmentValues] $VerticalAlignment # Property   ClosedXML.Excel.XLFontVerticalTextAlignmentValues VerticalAlignment {get;set;}
    )
    if ($Script:OfficeTrackerExcel) {
        $Worksheet = $Script:OfficeTrackerExcel['WorkSheet']
    } elseif (-not $Worksheet) {
        return
    }

    # Formatting of numbers/dates
    if ($Format) {
        $Worksheet.Cell($Row, $Column).Style.NumberFormat.Format = $Format
    } elseif ($FormatID) {
        $Worksheet.Cell($Row, $Column).Style.NumberFormat.NumberFormatID = $FormatID
    }

    if ($FontColor) {
        $ColorConverted = [ClosedXML.Excel.XLColor]::FromHtml((ConvertFrom-Color -Color $FontColor))
        $Worksheet.Cell($Row, $Column).Style.Font.FontColor = $ColorConverted
    }
    if ($null -ne $Bold) {
        $Worksheet.Cell($Row, $Column).Style.Font.Bold = $Bold
    }
    if ($null -ne $Italic) {
        $Worksheet.Cell($Row, $Column).Style.Font.Italic = $Italic
    }
    if ($null -ne $Strikethrough) {
        $Worksheet.Cell($Row, $Column).Style.Font.Strikethrough = $Strikethrough
    }
    if ($null -ne $Shadow) {
        $Worksheet.Cell($Row, $Column).Style.Font.Shadow = $Shadow
    }
    if ($FontSize) {
        $Worksheet.Cell($Row, $Column).Style.Font.FontSize = $FontSize
    }
    if ($null -ne $Underline) {
        $Worksheet.Cell($Row, $Column).Style.Font.Underline = $Underline
    }
    if ($null -ne $VerticalAlignment) {
        $Worksheet.Cell($Row, $Column).Style.Font.VerticalAlignment = $VerticalAlignment
    }
    if ($null -ne $FontFamilyNumbering) {
        $Worksheet.Cell($Row, $Column).Style.Font.FontFamilyNumbering = $FontFamilyNumbering
    }
    if ($null -ne $FontCharSet) {
        $Worksheet.Cell($Row, $Column).Style.Font.FontCharSet = $FontCharSet
    }

    <# $Worksheet.Cell($Row, $Column).Style
    Font               : False-False-None-False-Baseline-False-11-FF000000-Calibri-Swiss
    Alignment          : General-Bottom-0-False-ContextDependent-0-False-0-False-
    Border             : None-FF000000-None-FF000000-None-FF000000-None-FF000000-None-FF000000-False-False
    Fill               : None
    IncludeQuotePrefix : False
    NumberFormat       : 0-
    Protection         : Locked
    DateFormat         : 0-
    #>

    # $Worksheet.Cell($Row, $Column).Style.Fill | fl
    # BackgroundColor : Color Index: 64
    # PatternColor    : Color Index: 64
    # PatternType     : None

    if ($BackGroundColor) {
        $ColorConverted = [ClosedXML.Excel.XLColor]::FromHtml((ConvertFrom-Color -Color $BackGroundColor))
        $Worksheet.Cell($Row, $Column).Style.Fill.BackgroundColor = $ColorConverted
    }
    if ($PatternType) {
        $Worksheet.Cell($Row, $Column).Style.Fill.PatternType = $PatternType
    }

}

Register-ArgumentCompleter -CommandName Set-OfficeExcelValueStyle -ParameterName FontColor -ScriptBlock $Script:ScriptBlockColors
Register-ArgumentCompleter -CommandName Set-OfficeExcelValueStyle -ParameterName BackGroundColor -ScriptBlock $Script:ScriptBlockColors