---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Rename-OfficePowerPointSection
## SYNOPSIS
Renames a PowerPoint section.

## SYNTAX
```powershell
Rename-OfficePowerPointSection [[-Presentation] <PowerPointPresentation>] -Name <string> -NewName <string> [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Renames the first matching section in the presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Rename-OfficePowerPointSection -Presentation $ppt -Name 'Results' -NewName 'Deep Dive'
```

Renames the Results section to Deep Dive.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for the existing section name.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Existing section name.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NewName
New section name.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the renamed section.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to update.

```yaml
Type: PowerPointPresentation
Parameter Sets: (All)
Aliases: None
Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`
- `System.Boolean`

## RELATED LINKS

- None
