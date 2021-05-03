function New-OfficeWordTable {
    [cmdletBinding()]
    param(
        [Array] $DataTable,
        [DocumentFormat.OpenXml.Wordprocessing.TableLayoutValues] $TableLayout
    )


    if ($DataTable[0] -is [System.Collections.IDictionary]) {
        $Properties = 'Name', 'Value'
    } else {
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties -Property $IncludeProperty -ExcludeProperty $ExcludeProperty
    }

    $Table = [DocumentFormat.OpenXml.Wordprocessing.Table]::new()

    $TableProperties = [DocumentFormat.OpenXml.Wordprocessing.TableProperties]::new()


    #TextWrappingValues tableTextWrapping = TextWrappingValues.Around;

    $TableProperties.TableLayout = [DocumentFormat.OpenXml.Wordprocessing.TableLayout] @{
        Type = $TableLayout #[DocumentFormat.OpenXml.Wordprocessing.TableLayoutValues]::Autofit
    }
    # $TableProperties.TableWidth = [DocumentFormat.OpenXml.Wordprocessing.TableWidth] @{
    # Width = "0"
    #    Type = [DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues]::Auto
    # }


    $TextWrapping = [DocumentFormat.OpenXml.Wordprocessing.TextWrappingValues]::Auto
    #$tableProperties.Append($TextWrapping);
    #$tableProperties.Append($TableLook);
    #$TableProperties.Append($TableStyle);
    $TableProperties.TableStyle = New-OfficeWordTableStyle
    # $TableProperties.TableLook = New-OfficeWordTableLook
    #$TableProperties.TableBorders = New-OfficeWordTableBorder




    $table.Append($TableProperties)


    if (-not $SkipHeader) {
        $TableRow = [DocumentFormat.OpenXml.Wordprocessing.TableRow]::new()
        foreach ($Property in $Properties) {
            $TableCell = [DocumentFormat.OpenXml.Wordprocessing.TableCell]::new()
            $Paragraph = [DocumentFormat.OpenXml.Wordprocessing.Paragraph]::new()
            $TextProperty = New-OfficeWordText -Paragraph $Paragraph -Text $Property -ReturnObject


            $TableCellProperty = [DocumentFormat.OpenXml.Wordprocessing.TableCellProperties]::new()

            $TableCellWidth = [DocumentFormat.OpenXml.Wordprocessing.TableCellWidth]::new()
            $TableCellWidth.Width = 2400
            #$TableCellWidth.Type = [DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues]::Auto
            $TableCellProperty.TableCellWidth = $TableCellWidth
            $TableCellWidth.Type = [DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues]::Dxa
            # new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto })

            # $TableCell.Append($TableCellProperty)
            $TableCell.Append($TextProperty)
            $TableRow.Append($TableCell)

        }
        $Table.Append($TableRow)
    }

    # Table content
    if ($DataTable[0] -is [System.Collections.IDictionary]) {

    } elseif ($Properties -eq '*') {


    } else {
        # PSCustomObject
        foreach ($Data in $DataTable) {
            $TableRow = [DocumentFormat.OpenXml.Wordprocessing.TableRow]::new()

            foreach ($Property in $Properties) {
                $TableCell = [DocumentFormat.OpenXml.Wordprocessing.TableCell]::new()
                $Paragraph = [DocumentFormat.OpenXml.Wordprocessing.Paragraph]::new()
                if ($Data.$Property) {
                    $Text1 = New-OfficeWordText -Paragraph $Paragraph -Text $Data.$Property -ReturnObject
                }
                $TableCell.Append($Paragraph)
                $TableRow.Append($TableCell)
            }
            $Table.Append($TableRow)

        }
    }
    $null = $Document.MainDocumentPart.Document.Body.Append($Table)


}