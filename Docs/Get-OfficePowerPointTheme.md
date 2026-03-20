---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointTheme
## SYNOPSIS
Gets theme information for a PowerPoint presentation master.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointTheme [-Presentation <PowerPointPresentation>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Gets theme information for a PowerPoint presentation master.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointTheme -Presentation $ppt
```

Returns the theme name, theme colors, and configured fonts for master 0.

## PARAMETERS

### -Master
Slide master index to inspect.

```yaml
Type: Int32
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
Presentation to inspect (optional inside New-OfficePowerPoint).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointThemeInfo`

## RELATED LINKS

- None

