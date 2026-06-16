---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeVisioShapeLayout
## SYNOPSIS
Applies OfficeIMO Visio selection layout and layer operations to shapes.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeVisioShapeLayout [[-InputObject] <Object>] [-Page <VisioPage>] [-ShapeId <string[]>] [-Layer <string>] [-AlignHorizontal <VisioHorizontalAlignment>] [-AlignVertical <VisioVerticalAlignment>] [-Distribute <VisioDistributionAxis>] [-Grid] [-HorizontalStack] [-VerticalStack] [-Columns <int>] [-HorizontalSpacing <double>] [-VerticalSpacing <double>] [-PreserveFirstShapeCenter] [-NoRouteInternalConnectors] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Applies OfficeIMO Visio selection layout and layer operations to shapes.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficeVisioShapeLayout -AlignHorizontal 'Value'
```


## PARAMETERS

### -AlignHorizontal
Horizontal alignment inside the selected shapes' bounds.

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

### -AlignVertical
Vertical alignment inside the selected shapes' bounds.

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

### -Columns
Grid column count. Zero lets OfficeIMO choose a near-square grid.

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

### -Distribute
Distribute selected shapes along an axis.

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

### -Grid
Lay out selected shapes as a grid.

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

### -HorizontalSpacing
Horizontal spacing in inches for grid/stack layout.

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

### -HorizontalStack
Lay out selected shapes as a horizontal stack.

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

### -InputObject
Shapes, shape selections, or shape keys/ids to arrange.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Layer
Add selected shapes to this layer.

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

### -NoRouteInternalConnectors
Do not reroute internal connectors during OfficeIMO relayout.

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

### -Page
Page that owns the shapes. Optional inside New-OfficeVisio/VisioPage DSL scopes.

```yaml
Type: VisioPage
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
Emit arranged shapes.

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

### -PreserveFirstShapeCenter
Use the first selected shape as the grid origin instead of preserving the selection top-left.

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

### -ShapeId
Shape keys or ids to resolve on the target page.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -VerticalSpacing
Vertical spacing in inches for grid/stack layout.

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

### -VerticalStack
Lay out selected shapes as a vertical stack.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `OfficeIMO.Visio.VisioShape`

## RELATED LINKS

- None
