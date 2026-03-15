---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelUrlLinksByHeader
## SYNOPSIS
Converts cells under a header into external URL hyperlinks.

## SYNTAX
### ContextUsedRange (Default)
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextTable
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -TableName <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### ContextRange
```powershell
Set-OfficeExcelUrlLinksByHeader [-Header] <string> -Range <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Finds a column by header text and turns its non-empty values into external hyperlinks. You can target the worksheet used range, a named table, or a specific A1 range whose first row contains headers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }
```

Uses the used-range header row to find the `RFC` column and convert its values into links.

### EXAMPLE 2
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinksByHeader -Header 'RFC' -TableName 'LinksTable' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } -TitleScript { param($text) "Open $text" } }
```

Uses the `RFC` column inside `LinksTable` and controls the displayed link text.
