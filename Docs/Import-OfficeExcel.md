# Import-OfficeExcel

Imports rows from an Excel workbook as PowerShell objects using the OfficeIMO reader pipeline.

## Synopsis

```powershell
Import-OfficeExcel -Path .\Report.xlsx -WorksheetName Data
```

## Description

`Import-OfficeExcel` reads the used range from a worksheet by default and emits one object per row. Use `-Range` for an explicit A1 range or provide coordinate bounds with `-StartRow`, `-EndRow`, `-StartColumn`, and `-EndColumn`.

The first row is treated as headers unless `-NoHeader` is specified.

## Examples

```powershell
Import-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data
```

```powershell
Import-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data -Range 'A1:D20' -AsHashtable
```

```powershell
Import-OfficeExcel -Path .\Sales.xlsx -WorksheetName Data -StartRow 1 -EndRow 20 -StartColumn 1 -EndColumn 4 -AsDataTable
```
