---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeOpenDocument
## SYNOPSIS
Loads a native ODT, ODS, or ODP document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeOpenDocument [-Path] <string> [-Options <OdfLoadOptions>] [<CommonParameters>]
```

## DESCRIPTION
Loads a native ODT, ODS, or ODP document.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeOpenDocument -Path 'C:\Path'
```


## PARAMETERS

### -Options
Optional bounded package and XML settings.

```yaml
Type: OdfLoadOptions
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
Path to an ODT, ODS, or ODP file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String`

## OUTPUTS

- `OfficeIMO.OpenDocument.OdfDocument
OfficeIMO.OpenDocument.OdtDocument
OfficeIMO.OpenDocument.OdsDocument
OfficeIMO.OpenDocument.OdpPresentation`

## RELATED LINKS

- None
