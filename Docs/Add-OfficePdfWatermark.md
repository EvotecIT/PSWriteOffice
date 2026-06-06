---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfWatermark
## SYNOPSIS
Adds a generated-document text watermark.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfWatermark [-Text] <string> [-FontSize <double>] [-Opacity <double>] [-RotationAngle <double>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfWatermark [-Text] <string> -Document <PdfDocument> [-FontSize <double>] [-Opacity <double>] [-RotationAngle <double>] [-Color <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a generated-document text watermark.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfWatermark -Color 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePdfWatermark -Document 'Value'
```


## PARAMETERS

### -Color
Optional watermark color in #RRGGBB format.

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

### -Document
PDF document to update outside the DSL context.

```yaml
Type: PdfDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -FontSize
Watermark font size.

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

### -Opacity
Watermark opacity, 0 through 1.

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
Emit the updated document.

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

### -RotationAngle
Watermark rotation angle.

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

### -Text
Watermark text.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
