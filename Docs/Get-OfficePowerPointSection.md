---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointSection
## SYNOPSIS
Gets sections from a PowerPoint presentation.

## SYNTAX
```powershell
Get-OfficePowerPointSection [[-Presentation] <PowerPointPresentation>] [-Name <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Returns section metadata from an OfficeIMO PowerPoint presentation, including the section name, id, and zero-based slide indexes contained in the section.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointSection -Presentation $ppt
```

Lists all sections in the deck.

### EXAMPLE 2
```powershell
PS>Get-OfficePowerPointSection -Presentation $ppt -Name 'Intro'
```

Returns the section named Intro.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for `-Name`.

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
Optional section name filter.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to inspect.

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

## RELATED LINKS

- None
