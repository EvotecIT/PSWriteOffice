---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelSmartHyperlink
## SYNOPSIS
Sets an external hyperlink using a smart display strategy.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelSmartHyperlink [-Url] <string> [-Row <int>] [-Column <int>] [-Address <string>] [-Title <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelSmartHyperlink [-Url] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Row <int>] [-Column <int>] [-Address <string>] [-Title <string>] [-NoStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates an external hyperlink using OfficeIMO's smart display logic. If you omit `-Title`, the display text is inferred from the URL, for example `RFC 7208` for RFC links or the host name for normal URLs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Set-OfficeExcelSmartHyperlink -Address 'A2' -Url 'https://datatracker.ietf.org/doc/html/rfc7208' }
```

Creates a hyperlink that displays RFC 7208.

## PARAMETERS

### -Url
External URL to link to.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Nullable`1
Parameter Sets: Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Row
1-based row index.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Column
1-based column index.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Address
A1-style cell address.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Optional preferred display text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoStyle
Skip hyperlink styling (blue + underline).

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the worksheet after setting the link.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- [Set-OfficeExcelHyperlink](Set-OfficeExcelHyperlink.md)
- [Set-OfficeExcelHostHyperlink](Set-OfficeExcelHostHyperlink.md)
