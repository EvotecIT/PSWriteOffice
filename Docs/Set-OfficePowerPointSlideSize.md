---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideSize
## SYNOPSIS
Sets the slide size for a PowerPoint presentation.

## SYNTAX
### Preset (Default)
```powershell
Set-OfficePowerPointSlideSize [[-Presentation] <PowerPointPresentation>] -Preset <PowerPointSlideSizePreset> [-Portrait] [<CommonParameters>]
```

### Centimeters
```powershell
Set-OfficePowerPointSlideSize [[-Presentation] <PowerPointPresentation>] -WidthCm <double> -HeightCm <double> [<CommonParameters>]
```

### Inches
```powershell
Set-OfficePowerPointSlideSize [[-Presentation] <PowerPointPresentation>] -WidthInches <double> -HeightInches <double> [<CommonParameters>]
```

### Points
```powershell
Set-OfficePowerPointSlideSize [[-Presentation] <PowerPointPresentation>] -WidthPoints <double> -HeightPoints <double> [<CommonParameters>]
```

### Emus
```powershell
Set-OfficePowerPointSlideSize [[-Presentation] <PowerPointPresentation>] -WidthEmus <long> -HeightEmus <long> [<CommonParameters>]
```

## DESCRIPTION
Applies either a common PowerPoint size preset or a custom width and height to the whole presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointSlideSize -Presentation $ppt -Preset Screen16x9
```

Applies the standard 16:9 widescreen preset.

### EXAMPLE 2
```powershell
PS>Set-OfficePowerPointSlideSize -Presentation $ppt -Preset Screen4x3 -Portrait
```

Applies the 4:3 preset in portrait orientation.

### EXAMPLE 3
```powershell
PS>Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 25.4 -HeightCm 14.0
```

Applies a custom slide size of 25.4 by 14.0 centimeters.

## PARAMETERS

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

### -Preset
Preset slide size to apply.

```yaml
Type: PowerPointSlideSizePreset
Parameter Sets: Preset
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Portrait
Apply the preset in portrait orientation.

```yaml
Type: SwitchParameter
Parameter Sets: Preset
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthCm
Custom slide width in centimeters.

```yaml
Type: Double
Parameter Sets: Centimeters
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightCm
Custom slide height in centimeters.

```yaml
Type: Double
Parameter Sets: Centimeters
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthInches
Custom slide width in inches.

```yaml
Type: Double
Parameter Sets: Inches
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightInches
Custom slide height in inches.

```yaml
Type: Double
Parameter Sets: Inches
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthPoints
Custom slide width in points.

```yaml
Type: Double
Parameter Sets: Points
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightPoints
Custom slide height in points.

```yaml
Type: Double
Parameter Sets: Points
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthEmus
Custom slide width in EMUs.

```yaml
Type: Int64
Parameter Sets: Emus
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightEmus
Custom slide height in EMUs.

```yaml
Type: Int64
Parameter Sets: Emus
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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlideSize`

## RELATED LINKS

- [Set-OfficePowerPointSlideTransition](Set-OfficePowerPointSlideTransition.md)
