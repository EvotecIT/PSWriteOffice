---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficePowerPoint
## SYNOPSIS
Saves a presentation to disk.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficePowerPoint -Presentation <PowerPointPresentation> [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves a presentation to disk.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Save-OfficePowerPoint -Presentation $ppt -Show
```

Saves the current presentation and opens it in PowerPoint.

## PARAMETERS

### -Presentation
Presentation instance to save.

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Show
Launch the saved file in the default viewer.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

