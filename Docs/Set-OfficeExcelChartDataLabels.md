---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartDataLabels
## SYNOPSIS
Configures data labels and optional styling for an Excel chart.

## SYNTAX
```powershell
Set-OfficeExcelChartDataLabels [-Chart] <ExcelChart> [-ShowValue <bool>] [-ShowCategoryName <bool>] [-ShowSeriesName <bool>] [-ShowLegendKey <bool>] [-ShowPercent <bool>] [-Position <string>] [-NumberFormat <string>] [-SourceLinked <bool>] [-FontSizePoints <double>] [-Bold <bool>] [-Italic <bool>] [-Color <string>] [-FontName <string>] [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <double>] [-NoFill] [-NoLine] [<CommonParameters>]
```

## DESCRIPTION
Adds or updates chart data labels, then optionally applies text and shape styling to those labels.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$chart | Set-OfficeExcelChartDataLabels -ShowValue $true -Position OutsideEnd
```

Shows values as labels and positions them outside the series where supported.

### EXAMPLE 2
```powershell
PS>$chart | Set-OfficeExcelChartDataLabels -ShowValue $true -ShowPercent $true -NumberFormat '0.0%' -FillColor '#FFF2CC'
```

Shows values and percentages, applies a number format, and colors the label background.

## PARAMETERS

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: (All)
Aliases: None
Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -ShowValue
Show values in labels.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: True
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowCategoryName
Show category names in labels.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowSeriesName
Show series names in labels.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowLegendKey
Show legend keys in labels.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowPercent
Show percentages in labels.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: True
```

### -Position
Optional data label position.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

Valid values: `BestFit`, `Bottom`, `Center`, `InsideBase`, `InsideEnd`, `Left`, `OutsideEnd`, `Right`, `Top`

### -NumberFormat
Optional number format code.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourceLinked
Keep number formatting linked to the source cells.

```yaml
Type: Boolean
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: False
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontSizePoints
Optional label font size in points.

```yaml
Type: Nullable`1
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Bold
Optional bold setting for label text.

```yaml
Type: Nullable`1
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Italic
Optional italic setting for label text.

```yaml
Type: Nullable`1
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Color
Optional label text color in hex format.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FontName
Optional label font name.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FillColor
Optional label fill color in hex format.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineColor
Optional label line color in hex format.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineWidthPoints
Optional label border width in points.

```yaml
Type: Nullable`1
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoFill
Remove label fill.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoLine
Remove label border.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None
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

- [Set-OfficeExcelChartLegend](Set-OfficeExcelChartLegend.md)
- [Set-OfficeExcelChartStyle](Set-OfficeExcelChartStyle.md)
