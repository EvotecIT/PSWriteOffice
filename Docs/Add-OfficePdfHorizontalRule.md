---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfHorizontalRule
## SYNOPSIS
Adds a horizontal rule divider to a generated PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfHorizontalRule [-Thickness <double>] [-Color <string>] [-SpacingBefore <double>] [-SpacingAfter <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfHorizontalRule -Document <PdfDocument> [-Thickness <double>] [-Color <string>] [-SpacingBefore <double>] [-SpacingAfter <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a horizontal rule divider to a generated PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfDivider.pdf {
  Add-OfficePdfHeading -Text 'Executive summary'
  Add-OfficePdfParagraph -Text 'The service is healthy.'
  Add-OfficePdfHorizontalRule -Color '#CBD5E1' -Thickness 0.75 -SpacingBefore 10 -SpacingAfter 10
  Add-OfficePdfHeading -Text 'Signals' -Level 2
}
```

Adds a visual divider between generated PDF sections.

## PARAMETERS

### -Color
Rule color in #RRGGBB format.

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

### -SpacingAfter
Spacing after the rule in PDF points.

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

### -SpacingBefore
Spacing before the rule in PDF points.

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

### -Thickness
Rule thickness in PDF points.

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
