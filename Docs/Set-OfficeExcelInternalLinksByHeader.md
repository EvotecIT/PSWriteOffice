---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelInternalLinksByHeader
## SYNOPSIS
Converts cells under a header into internal workbook links.

## SYNTAX
### ContextUsedRange (Default)
```powershell
Set-OfficeExcelInternalLinksByHeader [-Header] <string> [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextTable
```powershell
Set-OfficeExcelInternalLinksByHeader [-Header] <string> -TableName <string> [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextRange
```powershell
Set-OfficeExcelInternalLinksByHeader [-Header] <string> -Range <string> [-DestinationSheetScript <scriptblock>] [-DisplayScript <scriptblock>] [-TargetAddress <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Finds a column by header text and converts the data cells under that header into internal workbook hyperlinks. You can scope the operation to the used range, a named table, or a specific A1 range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' }
```

Uses the used range header row to find the Sheet column and links each value to the matching sheet.

### EXAMPLE 2
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelInternalLinksByHeader -Header 'Sheet' -TableName 'SummaryTable' -DisplayScript { param($text) \"Open $text\" } }
```

Links the Sheet column inside SummaryTable and customizes the display text.
