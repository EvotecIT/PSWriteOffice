---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointSection
## SYNOPSIS
Adds a section to a PowerPoint presentation.

## SYNTAX
```powershell
Add-OfficePowerPointSection [[-Presentation] <PowerPointPresentation>] -Name <string> [-StartSlideIndex <int>] [<CommonParameters>]
```

## DESCRIPTION
Creates a section in a presentation starting at the requested zero-based slide index. Inside `New-OfficePowerPoint`, the current slide can provide the starting point automatically.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficePowerPointSection -Presentation $ppt -Name 'Results' -StartSlideIndex 2
```

Creates a Results section starting at slide index 2.

## PARAMETERS

### -Name
Name of the section to add.

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

### -StartSlideIndex
Zero-based slide index where the section should start.

```yaml
Type: Nullable`1
Parameter Sets: (All)
Aliases: None
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

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None
