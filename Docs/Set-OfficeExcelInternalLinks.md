---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelInternalLinks
## SYNOPSIS
Converts cells in a range into internal workbook links.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelInternalLinks [-Range] <string> [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelInternalLinks [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Turns each non-empty cell in the specified range into an internal hyperlink. By default, the cell text is used as both the destination sheet name and the display text.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinks -Range 'A2:A10' }
```

Links each value in A2:A10 to the sheet with the same name.

### EXAMPLE 2
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinks -Range 'A2:A10' -DisplayScript { param($text) \"Open $text\" } }
```

Links each value in A2:A10 and changes the displayed text.
