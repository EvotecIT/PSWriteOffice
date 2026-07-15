---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeAsciiDoc
## SYNOPSIS
Saves an OfficeIMO AsciiDoc document.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeAsciiDoc [-Document] <AsciiDocDocument> [-Path] <string> [-Options <AsciiDocWriterOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an OfficeIMO AsciiDoc document.

## EXAMPLES

### EXAMPLE 1
```powershell
Save-OfficeAsciiDoc -Path 'C:\Path'
```


## PARAMETERS

### -Document
AsciiDoc document to save.

```yaml
Type: AsciiDocDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Options
Optional writer settings.

```yaml
Type: AsciiDocWriterOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Return the saved document.

```yaml
Type: SwitchParameter
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
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.AsciiDoc.AsciiDocDocument`

## OUTPUTS

- `OfficeIMO.AsciiDoc.AsciiDocDocument`

## RELATED LINKS

- None
