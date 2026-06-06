---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePdfFooter
## SYNOPSIS
Sets running PDF footer text.

## SYNTAX
### Context (Default)
```powershell
Set-OfficePdfFooter [[-Text] <string>] [-Align <PdfAlign>] [-FontSize <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficePdfFooter [[-Text] <string>] -Document <PdfDocument> [-Align <PdfAlign>] [-FontSize <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets running PDF footer text.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePdf -Path .\Examples\Documents\PdfFooter.pdf {
                Set-OfficePdfFooter -Text 'Page {page} of {pages}' -Align Center -FontSize 8
                Add-OfficePdfHeading -Text 'Report with footer'
                Add-OfficePdfPageBreak
                Add-OfficePdfParagraph -Text 'The footer includes generated page numbers.'
            }
```

Uses page placeholders in a running footer.

## PARAMETERS

### -Align
Footer alignment.

```yaml
Type: PdfAlign
Parameter Sets: Context, Document
Aliases: None
Possible values: Left, Center, Right, Justify

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
Footer font size in PDF points.

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

### -Text
Footer text. Supports {page} and {pages}.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
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
