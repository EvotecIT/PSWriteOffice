---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointShapeLayout
## SYNOPSIS
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## SYNTAX
### Align (Default)
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Align <PowerPointShapeAlignment> [-Slide <PowerPointSlide>] [-ToSlide] [-MarginPoints <double>] [-PassThru] [<CommonParameters>]
```

### Distribute
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Distribute <PowerPointShapeDistribution> [-Slide <PowerPointSlide>] [-CrossAxisAlign <PowerPointShapeAlignment>] [-SpacingPoints <double>] [-ToSlide] [-MarginPoints <double>] [-Center] [-PassThru] [<CommonParameters>]
```

### Grid
```powershell
Set-OfficePowerPointShapeLayout [-InputObject] <Object> -Grid [-Slide <PowerPointSlide>] [-Columns <int>] [-Rows <int>] [-AutoGrid] [-GutterXPoints <double>] [-GutterYPoints <double>] [-Flow <PowerPointShapeGridFlow>] [-NoResize] [-ToSlide] [-MarginPoints <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Aligns, distributes, or arranges PowerPoint shapes using OfficeIMO layout helpers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Find-OfficePowerPointShape -Slide $slide -Name 'Kpi.*' |
                Set-OfficePowerPointShapeLayout -Align Top
```

Uses OfficeIMO.PowerPoint to align all matching shapes to the top edge of their selection bounds.

## PARAMETERS

### -Align
Alignment operation.

```yaml
Type: PowerPointShapeAlignment
Parameter Sets: Align
Aliases: None
Possible values: Left, Center, Right, Top, Middle, Bottom

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AutoGrid
Let OfficeIMO choose the grid dimensions.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Center
Center a fixed-spacing distribution within its bounds.

```yaml
Type: SwitchParameter
Parameter Sets: Distribute
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Columns
Grid column count. Omit with AutoGrid.

```yaml
Type: Nullable`1
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CrossAxisAlign
Optional cross-axis alignment for even distribution.

```yaml
Type: Nullable`1
Parameter Sets: Distribute
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Distribute
Distribution operation.

```yaml
Type: PowerPointShapeDistribution
Parameter Sets: Distribute
Aliases: None
Possible values: Horizontal, Vertical

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Flow
Fill the grid column-by-column instead of row-by-row.

```yaml
Type: PowerPointShapeGridFlow
Parameter Sets: Grid
Aliases: None
Possible values: RowMajor, ColumnMajor

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Grid
Arrange shapes in a grid.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GutterXPoints
Horizontal grid gutter in points.

```yaml
Type: Double
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -GutterYPoints
Vertical grid gutter in points.

```yaml
Type: Double
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
PowerPoint shapes or shape info records from Get-OfficePowerPointShape or Find-OfficePowerPointShape.

```yaml
Type: Object
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -MarginPoints
Use slide content bounds with the supplied margin in points.

```yaml
Type: Nullable`1
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoResize
Keep each shape's current size when arranging in a grid.

```yaml
Type: SwitchParameter
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the arranged shapes.

```yaml
Type: SwitchParameter
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Rows
Grid row count. Omit with AutoGrid.

```yaml
Type: Nullable`1
Parameter Sets: Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Slide that owns raw PowerPointShape inputs. Shape info records carry their own slide.

```yaml
Type: PowerPointSlide
Parameter Sets: Align, Distribute, Grid
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SpacingPoints
Fixed spacing between distributed shapes in points.

```yaml
Type: Nullable`1
Parameter Sets: Distribute
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ToSlide
Use the full slide bounds instead of the current selection bounds.

```yaml
Type: SwitchParameter
Parameter Sets: Align, Distribute, Grid
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

- `System.Object`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointShape`

## RELATED LINKS

- None
