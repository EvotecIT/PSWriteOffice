---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Close-OfficePowerPoint
## SYNOPSIS
Closes a PowerPoint presentation and optionally saves it.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficePowerPoint -Presentation <PowerPointPresentation> [-Save] [-Show] [-Password <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a cmdlet wrapper so PowerShell scripts do not need to call Dispose directly.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ppt = Get-OfficePowerPoint -FilePath .\deck.pptx; Close-OfficePowerPoint -Presentation $ppt
```

Releases the loaded presentation instance.

### EXAMPLE 2
```powershell
PS> Close-OfficePowerPoint -Presentation $ppt -Save -Show
```

Saves the presentation, opens it in PowerPoint, and releases the object.

## PARAMETERS

### -Password
Password used to save the presentation as an encrypted package.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to close.

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

### -Save
Persist changes before closing.

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

### -Show
Open the presentation in PowerPoint after saving.

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
