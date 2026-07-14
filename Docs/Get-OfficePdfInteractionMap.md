---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfInteractionMap
## SYNOPSIS
Builds text-selection and interactive hit regions for one PDF page.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePdfInteractionMap [-Path] <string> [-Page <int>] [-Options <PdfPageInteractionOptions>] [-ReadOptions <PdfReadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Builds text-selection and interactive hit regions for one PDF page.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficePdfInteractionMap -Path 'C:\Path'
```


## PARAMETERS

### -Options
Optional text-region limits.

```yaml
Type: PdfPageInteractionOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Page
One-based page number.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Source PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReadOptions
Optional bounded PDF parsing settings.

```yaml
Type: PdfReadOptions
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfPageInteractionMap`

## RELATED LINKS

- None
