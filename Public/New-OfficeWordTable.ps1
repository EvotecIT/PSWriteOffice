function New-OfficeWordTable {
    [cmdletBinding()]
    param(
        [OfficeIMO.Word.WordDocument] $Document,
        [Array] $DataTable,
        [OfficeIMO.Word.WordTableStyle] $Style = [OfficeIMO.Word.WordTableStyle]::TableGrid,
        [string] $TableLayout,
        [switch] $SkipHeader,
        [switch] $Suppress
    )

    if (-not $Document) {
        Write-Warning -Message "New-OfficeWordTable - Document is not specified. Please provide valid document."
        return
    }

    if ($DataTable[0] -is [System.Collections.IDictionary]) {
        $Properties = 'Name', 'Value'
    } else {
        $Properties = Select-Properties -Objects $DataTable -AllProperties:$AllProperties -Property $IncludeProperty -ExcludeProperty $ExcludeProperty
    }
    $CountRows = 0
    $CountColumns = 0

    $RowsCount = $DataTable.Count
    $ColumnsCount = $Properties.Count

    if (-not $SkipHeader) {
        # Since we need header we add additional row
        $Table = $Document.AddTable($RowsCount + 1, $ColumnsCount, $Style)
        # Add table header, if we don't explicitly ask for it to be skipped
        foreach ($Property in $Properties) {
            $Table.Rows[0].Cells[$CountColumns].Paragraphs[0].Text = $Property
            $CountColumns += 1
        }
        $CountRows += 1
    } else {
        # No header so less rows
        $Table = $Document.AddTable($RowsCount, $ColumnsCount, $Style)
    }

    # add table data
    foreach ($Row in $DataTable) {
        $CountColumns = 0
        foreach ($P in $Properties) {
            $Table.Rows[$CountRows].Cells[$CountColumns].Paragraphs[0].Text = $Row.$P
            $CountColumns += 1
        }
        $CountRows += 1
    }

    # return table object
    if (-not $Suppress) {
        $Table
    }

    <#
    # Table content
    if ($DataTable[0] -is [System.Collections.IDictionary]) {

    } elseif ($Properties -eq '*') {


    } else {
        # PSCustomObject
        foreach ($Data in $DataTable) {
            #$TableRow = [DocumentFormat.OpenXml.Wordprocessing.TableRow]::new()

            foreach ($Property in $Properties) {
                $TableCell = [DocumentFormat.OpenXml.Wordprocessing.TableCell]::new()
                $Paragraph = [DocumentFormat.OpenXml.Wordprocessing.Paragraph]::new()
                if ($Data.$Property) {
                   # $Text1 = New-OfficeWordText -Paragraph $Paragraph -Text $Data.$Property -ReturnObject
                }
                $TableCell.Append($Paragraph)
                $TableRow.Append($TableCell)
            }
            #$Table.Append($TableRow)

        }
    }
    #>

}