---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordShape
## SYNOPSIS
Adds a basic OfficeIMO Word shape to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordShape [-Type <ShapeType>] [-Width <double>] [-Height <double>] [-Left <double>] [-Top <double>] [-FillColor <string>] [-StrokeColor <string>] [-StrokeWidth <double>] [-Title <string>] [-Description <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a basic OfficeIMO Word shape to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficeWordShape -Description 'Value'
```


## PARAMETERS

### -Description
Optional alternate text metadata.

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

### -FillColor
Fill color as #RRGGBB.

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
Height in points.

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

### -Left
Anchored left position in points. When omitted, the shape is inline.

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

### -PassThru
Emit the created shape.

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

### -StrokeColor
Stroke color as #RRGGBB.

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

### -StrokeWidth
Stroke width in points.

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

### -Title
Optional title metadata.

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

### -Top
Anchored top position in points. When omitted, the shape is inline.

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

### -Type
Shape type to add.

```yaml
Type: ShapeType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Rectangle, Ellipse, Line, RoundedRectangle, Triangle, Diamond, Pentagon, Hexagon, RightArrow, LeftArrow, UpArrow, DownArrow, Star5, Heart, Cloud, Donut, Can, Cube

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Width in points.

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

- `None`

## OUTPUTS

- `OfficeIMO.Word.WordShape`

## RELATED LINKS

- None
