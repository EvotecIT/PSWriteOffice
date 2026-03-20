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
Set-OfficePowerPointSlideSize -Preset <PowerPointSlideSizePreset> [-Presentation <PowerPointPresentation>] [-Portrait] [<CommonParameters>]
```

### Centimeters
```powershell
Set-OfficePowerPointSlideSize -WidthCm <double> -HeightCm <double> [-Presentation <PowerPointPresentation>] [<CommonParameters>]
```

### Inches
```powershell
Set-OfficePowerPointSlideSize -WidthInches <double> -HeightInches <double> [-Presentation <PowerPointPresentation>] [<CommonParameters>]
```

### Points
```powershell
Set-OfficePowerPointSlideSize -WidthPoints <double> -HeightPoints <double> [-Presentation <PowerPointPresentation>] [<CommonParameters>]
```

### Emus
```powershell
Set-OfficePowerPointSlideSize -WidthEmus <long> -HeightEmus <long> [-Presentation <PowerPointPresentation>] [<CommonParameters>]
```

## DESCRIPTION
Sets the slide size for a PowerPoint presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointSlideSize -Presentation $ppt -Preset Screen16x9
```

Applies the 16:9 widescreen preset to the presentation.

### EXAMPLE 2
```powershell
PS>Set-OfficePowerPointSlideSize -Presentation $ppt -WidthCm 25.4 -HeightCm 14.0
```

Sets the presentation slide size to a custom 25.4 x 14.0 cm layout.

## PARAMETERS

### -HeightCm
Custom slide height in centimeters.

```yaml
Type: Double
Parameter Sets: Centimeters
Aliases: None
Possible values: 

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
Possible values: 

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
Possible values: 

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
Possible values: 

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
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Presentation
Presentation to update (optional inside New-OfficePowerPoint).

```yaml
Type: PowerPointPresentation
Parameter Sets: Preset, Centimeters, Inches, Points, Emus
Aliases: None
Possible values: 

Required: False
Position: named
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
Possible values: Screen4x3, Screen16x9, Screen16x10

Required: True
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
Possible values: 

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
Possible values: 

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
Possible values: 

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

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlideSize`

## RELATED LINKS

- None

