---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartTrendline
## SYNOPSIS
Adds or replaces an Excel chart series trendline.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartTrendline -Chart <ExcelChart> -SeriesIndex <int> -Type <string> [-Order <int>] [-Period <int>] [-Forward <double>] [-Backward <double>] [-Intercept <double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <string>] [-LineWidthPoints <double>] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartTrendline -Chart <ExcelChart> -SeriesName <string> -Type <string> [-IgnoreCase <bool>] [-Order <int>] [-Period <int>] [-Forward <double>] [-Backward <double>] [-Intercept <double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <string>] [-LineWidthPoints <double>] [<CommonParameters>]
```

## DESCRIPTION
Adds or replaces an Excel chart series trendline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartTrendline -SeriesIndex 0 -Type Polynomial -Order 2 -DisplayEquation -DisplayRSquared
```

Adds a polynomial trendline to the first series.

## PARAMETERS

### -Backward
Backward forecast units.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
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
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -DisplayEquation
Display the trendline equation.

```yaml
Type: SwitchParameter
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DisplayRSquared
Display the R-squared value.

```yaml
Type: SwitchParameter
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Forward
Forward forecast units.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IgnoreCase
Ignore case when matching series name.

```yaml
Type: Boolean
Parameter Sets: Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Intercept
Trendline intercept.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineColor
Trendline line color in hex format.

```yaml
Type: String
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LineWidthPoints
Trendline line width in points.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Order
Polynomial order.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Period
Moving-average period.

```yaml
Type: Nullable`1
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SeriesIndex
Zero-based series index.

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SeriesName
Series name.

```yaml
Type: String
Parameter Sets: Name
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Type
Trendline type.

```yaml
Type: String
Parameter Sets: Index, Name
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

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None
