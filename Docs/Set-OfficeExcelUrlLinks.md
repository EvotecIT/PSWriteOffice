---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelUrlLinks
## SYNOPSIS
Converts cells in a range into external URL hyperlinks.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelUrlLinks [-Range] <string> -UrlScript <scriptblock> [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelUrlLinks [-Range] <string> -UrlScript <scriptblock> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-TitleScript <scriptblock>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Turns each non-empty cell in the specified range into an external hyperlink. Use `-UrlScript` to map the existing cell text to a target URL, and optionally use `-TitleScript` to control the display text. When `-TitleScript` is omitted, PSWriteOffice uses OfficeIMO's smart display behavior.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinks -Range 'D2:D10' -UrlScript { param($text) "https://datatracker.ietf.org/doc/html/$text" } }
```

Links each RFC value in D2:D10 to its datatracker page and uses smart display text such as `RFC 7208`.

### EXAMPLE 2
```powershell
PS>ExcelSheet 'Summary' { Set-OfficeExcelUrlLinks -Range 'D2:D10' -UrlScript { param($text) "https://example.org/docs/$text" } -TitleScript { param($text) "Open $text" } }
```

Links each value in D2:D10 and replaces the displayed text.
