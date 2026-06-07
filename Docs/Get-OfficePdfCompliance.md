---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePdfCompliance
## SYNOPSIS
Gets a generated PDF document compliance readiness report.

## SYNTAX
### Context (Default)
```powershell
Get-OfficePdfCompliance [-Profile <PdfComplianceProfile>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficePdfCompliance -Document <PdfDocument> [-Profile <PdfComplianceProfile>] [<CommonParameters>]
```

## DESCRIPTION
Gets a generated PDF document compliance readiness report.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pdf = New-OfficePdf {
    Set-OfficePdfCompliance -Profile PdfA3B -Groundwork
    Add-OfficePdfHeading -Text 'Compliance readiness'
} -NoSave
$pdf | Get-OfficePdfCompliance -Profile PdfA3B
```

Returns the OfficeIMO.Pdf readiness report before saving.

## PARAMETERS

### -Document
Generated PDF document to assess outside the DSL context.

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

### -Profile
Compliance profile to assess. When omitted, the document's configured profile is used.

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

- `OfficeIMO.Pdf.PdfComplianceReadinessReport`

## RELATED LINKS

- None
