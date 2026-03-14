---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointBullets
## SYNOPSIS
Adds a bulleted list to a PowerPoint slide.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficePowerPointBullets [-Bullets] <string[]> [-Slide <PowerPointSlide>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Level <int>] [-BulletChar <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a bulleted list to a PowerPoint slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficePowerPointBullets -Slide $slide -Bullets 'Wins','Risks','Next Steps' -X 60 -Y 120 -Width 400 -Height 200
```

Creates a bullet list textbox.

## PARAMETERS

### -BulletChar
Optional bullet character (defaults to •).

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

### -Bullets
Bullet items to render.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Textbox height in points.

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

### -Level
List level (0-8).

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

### -Slide
Target slide that will receive the bullet list (optional inside DSL).

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
Textbox width in points.

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

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None

