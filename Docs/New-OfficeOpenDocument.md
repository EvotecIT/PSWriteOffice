---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficeOpenDocument
## SYNOPSIS
Creates a native ODT, ODS, or ODP document.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficeOpenDocument [-Kind] <OdfDocumentKind> [[-Path] <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Creates a native ODT, ODS, or ODP document.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeOpenDocument -Path 'C:\Path'
```


## PARAMETERS

### -Kind
OpenDocument text, spreadsheet, or presentation kind.

```yaml
Type: OdfDocumentKind
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Text, Spreadsheet, Presentation

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Optional initial destination path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfDocument`

## RELATED LINKS

- None
