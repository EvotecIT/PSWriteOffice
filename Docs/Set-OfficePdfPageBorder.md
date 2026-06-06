---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfPageBorder
## SYNOPSIS
Sets or clears the generated PDF page border decoration.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfPageBorder [-Color <string>] [-Width <double>] [-Inset <double>] [-Opacity <double>] [-DashStyle <OfficeStrokeDashStyle>] [-Clear] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfPageBorder -Document <PdfDocument> [-Color <string>] [-Width <double>] [-Inset <double>] [-Opacity <double>] [-DashStyle <OfficeStrokeDashStyle>] [-Clear] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets or clears the generated PDF page border decoration.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfPageBorder.pdf {
                Set-OfficePdfPageBorder -Color '#CBD5E1' -Width 0.75 -Inset 24 -Opacity 0.8
                Add-OfficePdfHeading -Text 'Bordered report'
            }
```

Decorates generated pages with an OfficeIMO.Pdf page border.

## PARAMETERS

### -Clear
Clear the generated PDF page border decoration.

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

### -Color
Border color in #RRGGBB format.

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

### -DashStyle
Border dash style.

```yaml
Type: OfficeStrokeDashStyle
Parameter Sets: Context, Document
Aliases: None
Possible values: Solid, Dash, Dot, DashDot

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

### -Inset
Distance from the page edge to the border path in PDF points.

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
Border opacity from 0 through 1.

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

### -Width
Border stroke width in PDF points.

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
