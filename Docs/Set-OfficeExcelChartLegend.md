---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartLegend
## SYNOPSIS
Configures legend visibility and styling for an Excel chart.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelChartLegend -Chart <ExcelChart> [-Position <string>] [-Overlay <bool>] [-Hide] [-FontSizePoints <double>] [-Bold <bool>] [-Italic <bool>] [-Color <string>] [-FontName <string>] [<CommonParameters>]
```

## DESCRIPTION
Configures legend visibility and styling for an Excel chart.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$chart | Set-OfficeExcelChartLegend -Position Right
```

Shows the legend on the right side of the chart and returns the chart for chaining.

## PARAMETERS

### -Bold
Optional bold setting for legend text.

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

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Color
Optional legend text color in hex format.

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

### -FontName
Optional legend font name.

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

### -FontSizePoints
Optional legend font size in points.

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

### -Hide
Hide the legend instead of positioning it.

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

### -Italic
Optional italic setting for legend text.

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

### -Overlay
Overlay the legend on the chart area.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Position
Legend position.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Bottom, Left, Right, Top, TopRight

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None

