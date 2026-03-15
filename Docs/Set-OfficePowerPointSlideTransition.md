---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideTransition
## SYNOPSIS
Sets the transition used when advancing to a slide.

## SYNTAX
```powershell
Set-OfficePowerPointSlideTransition [[-Slide] <PowerPointSlide>] -Transition <SlideTransition> [<CommonParameters>]
```

## DESCRIPTION
Applies one of the `OfficeIMO.PowerPoint` slide transition values to a slide object or to the current slide inside the PowerPoint DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTransition -Transition Fade
```

Applies a fade transition to the first slide.

### EXAMPLE 2
```powershell
PS>New-OfficePowerPoint -Path .\deck.pptx { PptSlide { PptTitle -Title 'Status'; PptTransition -Transition Morph } }
```

Applies a morph transition inside the PowerPoint DSL.

## PARAMETERS

### -Slide
Slide to update.

```yaml
Type: PowerPointSlide
Parameter Sets: (All)
Aliases: None
Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Transition
Transition to apply.

```yaml
Type: SlideTransition
Parameter Sets: (All)
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- [Set-OfficePowerPointSlideSize](Set-OfficePowerPointSlideSize.md)
