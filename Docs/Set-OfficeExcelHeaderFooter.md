---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelHeaderFooter
## SYNOPSIS
Sets worksheet header and footer text and optional images.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelHeaderFooter [-HeaderLeft <string>] [-HeaderCenter <string>] [-HeaderRight <string>] [-FooterLeft <string>] [-FooterCenter <string>] [-FooterRight <string>] [-DifferentFirstPage] [-DifferentOddEven] [-AlignWithMargins <bool>] [-ScaleWithDocument <bool>] [-HeaderImagePath <string>] [-HeaderImageUrl <string>] [-HeaderImagePosition <HeaderFooterPosition>] [-FooterImagePath <string>] [-FooterImageUrl <string>] [-FooterImagePosition <HeaderFooterPosition>] [-ImageWidthPoints <double>] [-ImageHeightPoints <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelHeaderFooter -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-HeaderLeft <string>] [-HeaderCenter <string>] [-HeaderRight <string>] [-FooterLeft <string>] [-FooterCenter <string>] [-FooterRight <string>] [-DifferentFirstPage] [-DifferentOddEven] [-AlignWithMargins <bool>] [-ScaleWithDocument <bool>] [-HeaderImagePath <string>] [-HeaderImageUrl <string>] [-HeaderImagePosition <HeaderFooterPosition>] [-FooterImagePath <string>] [-FooterImageUrl <string>] [-FooterImagePosition <HeaderFooterPosition>] [-ImageWidthPoints <double>] [-ImageHeightPoints <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets worksheet header and footer text and optional images.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Set-OfficeExcelHeaderFooter -HeaderCenter 'Demo' -FooterRight 'Page &P of &N' }
```

Applies header and footer text to the worksheet.

## PARAMETERS

### -AlignWithMargins
Align header/footer with margins (default: true).

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DifferentFirstPage
Use a different header/footer on the first page.

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

### -DifferentOddEven
Use different headers/footers on odd/even pages.

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

### -FooterCenter
Center footer text.

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

### -FooterImagePath
Footer image file path.

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

### -FooterImagePosition
Footer image position.

```yaml
Type: HeaderFooterPosition
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FooterImageUrl
Footer image URL.

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

### -FooterLeft
Left footer text.

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

### -FooterRight
Right footer text.

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

### -HeaderCenter
Center header text.

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

### -HeaderImagePath
Header image file path.

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

### -HeaderImagePosition
Header image position.

```yaml
Type: HeaderFooterPosition
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeaderImageUrl
Header image URL.

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

### -HeaderLeft
Left header text.

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

### -HeaderRight
Right header text.

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

### -ImageHeightPoints
Image height in points.

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

### -ImageWidthPoints
Image width in points.

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

### -PassThru
Emit the worksheet after updating.

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

### -ScaleWithDocument
Scale header/footer with document (default: true).

```yaml
Type: Boolean
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

