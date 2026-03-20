---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelImageFromUrl
## SYNOPSIS
Adds an image from a URL anchored to a worksheet cell.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelImageFromUrl [-Url] <string> [-Row <int>] [-Column <int>] [-Address <string>] [-WidthPixels <int>] [-HeightPixels <int>] [-OffsetXPixels <int>] [-OffsetYPixels <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelImageFromUrl [-Url] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Row <int>] [-Column <int>] [-Address <string>] [-WidthPixels <int>] [-HeightPixels <int>] [-OffsetXPixels <int>] [-OffsetYPixels <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an image from a URL anchored to a worksheet cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelImageFromUrl -Address 'B2' -Url 'https://example.org/logo.png' -WidthPixels 120 -HeightPixels 40 }
```

Downloads the remote image and anchors it to cell B2.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

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
Possible values: 

Required: False
Position: named
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
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -HeightPixels
Image height in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OffsetXPixels
Horizontal offset in pixels from the cell origin.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OffsetYPixels
Vertical offset in pixels from the cell origin.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the worksheet after inserting the image.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values: 

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
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values: 

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
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Url
Image URL to download.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthPixels
Image width in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

