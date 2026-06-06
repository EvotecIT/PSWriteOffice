---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePdfTable
## SYNOPSIS
Adds a table to a PDF document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficePdfTable [-InputObject] <Object[]> [-Property <string[]>] [-Header <string[]>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficePdfTable [-InputObject] <Object[]> -Document <PdfDocument> [-Property <string[]>] [-Header <string[]>] [-Align <PdfAlign>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a table to a PDF document.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfTable -Align 'Value'
```


### EXAMPLE 2
```powershell
Add-OfficePdfTable -Document 'Value'
```


## PARAMETERS

### -Align
Table alignment.

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

### -Header
Header labels. Defaults to property names.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Objects or row arrays to render as a table.

```yaml
Type: Object[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
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

### -Property
Specific object properties to include.

```yaml
Type: String[]
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
