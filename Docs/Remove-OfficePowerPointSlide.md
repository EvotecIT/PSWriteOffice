---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Remove-OfficePowerPointSlide
## SYNOPSIS
Removes a slide by index.

## SYNTAX
### __AllParameterSets
```powershell
Remove-OfficePowerPointSlide -Presentation <PowerPointPresentation> -Index <int> [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes a slide by index.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Remove-OfficePowerPointSlide -Presentation $ppt -Index 0
```

Removes slide 1 from the deck.

## PARAMETERS

### -Index
Zero-based slide index.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to modify.

```yaml
Type: PowerPointPresentation
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

