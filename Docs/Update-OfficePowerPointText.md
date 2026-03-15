---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficePowerPointText
## SYNOPSIS
Replaces text in a PowerPoint slide or presentation.

## SYNTAX
### Auto (Default)
```powershell
Update-OfficePowerPointText -OldValue <string> -NewValue <string> [-IncludeTables <bool>] [-IncludeNotes] [<CommonParameters>]
```

### Presentation
```powershell
Update-OfficePowerPointText [-Presentation] <PowerPointPresentation> -OldValue <string> -NewValue <string> [-IncludeTables <bool>] [-IncludeNotes] [<CommonParameters>]
```

### Slide
```powershell
Update-OfficePowerPointText [-Slide] <PowerPointSlide> -OldValue <string> -NewValue <string> [-IncludeTables <bool>] [-IncludeNotes] [<CommonParameters>]
```

## DESCRIPTION
Replaces matching text in a slide or across the whole presentation. The alias `Replace-OfficePowerPointText` is also exported for convenience.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Update-OfficePowerPointText -Presentation $ppt -OldValue 'FY24' -NewValue 'FY25' -IncludeNotes
```

Replaces FY24 with FY25 across the deck, including notes.

### EXAMPLE 2
```powershell
PS>Replace-OfficePowerPointText -Slide $slide -OldValue 'Draft' -NewValue 'Final'
```

Replaces Draft with Final in a single slide.

## PARAMETERS

### -IncludeNotes
Include notes text in the replacement operation.

```yaml
Type: SwitchParameter
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeTables
Include table cells in the replacement operation.

```yaml
Type: Boolean
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Required: False
Position: named
Default value: True
Accept pipeline input: False
Accept wildcard characters: True
```

### -NewValue
Replacement text.

```yaml
Type: String
Parameter Sets: Auto, Presentation, Slide
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OldValue
Text to find.

```yaml
Type: String
Parameter Sets: Auto, Presentation, Slide
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
Parameter Sets: Presentation
Aliases: None
Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Slide
Slide to update.

```yaml
Type: PowerPointSlide
Parameter Sets: Slide
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
- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `System.Int32`

## RELATED LINKS

- None
