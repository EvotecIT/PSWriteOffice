---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPageSetup
## SYNOPSIS
Sets PDF page size, orientation, and margins.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfPageSetup [-PageSize <string>] [-Width <double>] [-Height <double>] [-Landscape] [-Margin <double>] [-Left <double>] [-Top <double>] [-Right <double>] [-Bottom <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfPageSetup -Document <PdfDocument> [-PageSize <string>] [-Width <double>] [-Height <double>] [-Landscape] [-Margin <double>] [-Left <double>] [-Top <double>] [-Right <double>] [-Bottom <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets PDF page size, orientation, and margins.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePdfPageSetup -Bottom 'Value'
```


### EXAMPLE 2
```powershell
Set-OfficePdfPageSetup -Document 'Value'
```


## PARAMETERS

### -Bottom
Bottom margin in PDF points.

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

### -Height
Custom page height in PDF points when -PageSize Custom is used.

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

### -Landscape
Use landscape orientation.

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

### -Left
Left margin in PDF points.

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

### -Margin
Uniform margin in PDF points.

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

### -PageSize
Page size name: A4, A5, Letter, Legal, or Custom.

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

### -Right
Right margin in PDF points.

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

### -Top
Top margin in PDF points.

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

### -Width
Custom page width in PDF points when -PageSize Custom is used.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Pdf.PdfDocument`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None
