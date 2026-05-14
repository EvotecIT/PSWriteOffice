# Export-OfficeExcel

Exports PowerShell objects to an Excel workbook using an ImportExcel-style operator surface backed by OfficeIMO.

## Synopsis

```powershell
$rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Data -AutoFit -FreezeTopRow
```

## Description

`Export-OfficeExcel` is the quick path for object-to-workbook reporting. It writes pipeline input to a worksheet, creates an Excel table by default, can auto-fit columns, freeze the header row, add a title, append rows, clear an existing sheet, and optionally open the saved workbook.

Use the lower-level `New-OfficeExcel` DSL when you need full workbook composition.

## Examples

```powershell
$rows = @(
    [pscustomobject]@{ Region = 'NA'; Revenue = 100 }
    [pscustomobject]@{ Region = 'EMEA'; Revenue = 200 }
)

$rows | Export-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data -TableName Sales -AutoFit -FreezeTopRow -BoldTopRow
```

```powershell
$moreRows | Export-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data -TableName Sales -Append -AutoFit
```

```powershell
$rows | Export-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data -ClearSheet -Title 'Sales Export' -Show
```

```powershell
Import-Module PSParseHTML

ConvertFrom-HtmlTable -Path .\report.html -TableId 'results' -AsDataTable -IncludeLinkUrls |
    Export-OfficeExcel -Path .\HtmlTables.xlsx -WorksheetName Results -TableName Results -AutoFit -FreezeTopRow
```

`ConvertFrom-HtmlTable` comes from PSParseHTML/HtmlTinkerX. PSWriteOffice intentionally does not parse HTML itself; it consumes the `DataTable`, `DataSet`, reader, or objects produced by the upstream parser.

## Notes

`-Append` skips headers by default. When `-TableName` identifies an existing table, or the target sheet has exactly one table, PSWriteOffice uses OfficeIMO table append support when available so the Excel table range grows with the new rows. Older OfficeIMO builds fall back to writing raw rows after the used range.
