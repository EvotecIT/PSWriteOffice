---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficePdfSanitized
## SYNOPSIS
Removes or quarantines active PDF content and embedded payloads with post-save proof.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficePdfSanitized [-Path] <string> [-OutputPath] <string> [-Options <PdfSanitizationOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes or quarantines active PDF content and embedded payloads with post-save proof.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = ConvertTo-OfficePdfSanitized -Path .\Input.pdf -OutputPath .\Safe.pdf
```

Writes the proven full-rewrite result and returns findings, mutation plan, and quarantine data.

## PARAMETERS

### -Options
Allowed actions, URI schemes, embedded-file policy, and rich-media policy.

```yaml
Type: PdfSanitizationOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Destination PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfSanitizationResult`

## RELATED LINKS

- None
