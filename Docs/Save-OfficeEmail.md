---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeEmail
## SYNOPSIS
Saves an email document as EML, MSG, or TNEF with fidelity diagnostics.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeEmail [-Path] <string> -Document <EmailDocument> [-Format <EmailFileFormat>] [-Options <EmailWriterOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an email document as EML, MSG, or TNEF with fidelity diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeEmail -Document 'Value'
```


## PARAMETERS

### -Document
Email document to save.

```yaml
Type: EmailDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Format
Optional explicit output format. By default it is inferred from the filename.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Options
Optional preservation, projection, encoding, and output limits.

```yaml
Type: EmailWriterOptions
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
Destination path.

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

- `OfficeIMO.Email.EmailDocument`

## OUTPUTS

- `OfficeIMO.Email.EmailWriteResult`

## RELATED LINKS

- None
