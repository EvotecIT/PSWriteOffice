---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointShape
## SYNOPSIS
Adds a basic shape to a slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointShape [-Slide <PowerPointSlide>] [-ShapeType <string>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Name <string>] [-FillColor <string>] [-OutlineColor <string>] [-OutlineWidth <double>] [<CommonParameters>]
```

## DESCRIPTION
Creates an auto shape at the requested coordinates and applies optional fill and outline styling.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficePowerPoint -Path .\Examples\Documents\PowerPointShape.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1
    Add-OfficePowerPointShape -Slide $slide -ShapeType Rectangle -X 60 -Y 120 -Width 220 -Height 90 -FillColor '#DDEEFF' -OutlineColor '#2563EB' -OutlineWidth 1
    Add-OfficePowerPointTextBox -Slide $slide -Text 'Highlighted status' -X 80 -Y 145 -Width 180 -Height 32
}
```

Creates a styled rectangle and overlays a text box.

## PARAMETERS

### -FillColor
Fill color (hex or named color).

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

### -Height
Shape height in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Name
Optional name assigned to the shape.

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

### -OutlineColor
Outline color (hex or named color).

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

### -OutlineWidth
Outline width in points.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShapeType
Shape geometry preset name (e.g., Rectangle, Ellipse, Line).

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

### -Slide
Target slide that will receive the shape (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Width
Shape width in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
Left offset (in points) from the slide origin.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Top offset (in points) from the slide origin.

```yaml
Type: Double
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
