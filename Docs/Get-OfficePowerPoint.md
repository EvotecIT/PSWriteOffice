---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPoint
## SYNOPSIS
Loads an existing PowerPoint presentation.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPoint -FilePath <string> [<CommonParameters>]
```

## DESCRIPTION
Loads an existing PowerPoint presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$ppt = Get-OfficePowerPoint -FilePath .\Quarterly.pptx
```

Reads Quarterly.pptx and exposes the presentation object.

## PARAMETERS

### -FilePath
Path to the .pptx file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
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

- `System.Object`

## RELATED LINKS

- None

