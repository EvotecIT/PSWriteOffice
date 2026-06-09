---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficePowerPointChart
## SYNOPSIS
Adds a chart to a PowerPoint slide.

## SYNTAX
### Default (Default)
```powershell
Add-OfficePowerPointChart [-Slide <PowerPointSlide>] [-Type <PowerPointChartType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Title <string>] [<CommonParameters>]
```

### Categorical
```powershell
Add-OfficePowerPointChart -InputObject <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Slide <PowerPointSlide>] [-Type <PowerPointChartType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Title <string>] [<CommonParameters>]
```

### Scatter
```powershell
Add-OfficePowerPointChart -InputObject <Object[]> -XProperty <string> -YProperty <string[]> [-Slide <PowerPointSlide>] [-Type <PowerPointChartType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Supports default chart data or object-based category/series mappings for standard and scatter charts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = @(
    [pscustomobject]@{ Month = 'Jan'; Sales = 42; Profit = 9 }
    [pscustomobject]@{ Month = 'Feb'; Sales = 55; Profit = 13 }
)
New-OfficePowerPoint -Path .\Examples\Documents\PowerPointChart.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1
    Add-OfficePowerPointChart -Slide $slide -InputObject $rows -CategoryProperty Month -SeriesProperty Sales,Profit -Title 'Monthly performance'
}
```

Creates a clustered column chart using Month for categories and Sales/Profit as series.

### EXAMPLE 2
```powershell
PS> $rows = @(
    [pscustomobject]@{ Quarter = 1; Revenue = 20 }
    [pscustomobject]@{ Quarter = 2; Revenue = 34 }
)
New-OfficePowerPoint -Path .\Examples\Documents\PowerPointScatter.pptx {
    $slide = Add-OfficePowerPointSlide -Layout 1
    Add-OfficePowerPointChart -Slide $slide -Type Scatter -InputObject $rows -XProperty Quarter -YProperty Revenue -Title 'Revenue trend'
}
```

Creates a scatter chart using Quarter on the X axis and Revenue on the Y axis.

## PARAMETERS

### -CategoryProperty
Property name used for category labels on standard charts.

```yaml
Type: String
Parameter Sets: Categorical
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Height
Chart height in points.

```yaml
Type: Double
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Source objects used to build chart data.

```yaml
Type: Object[]
Parameter Sets: Categorical, Scatter
Aliases: Data
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SeriesProperty
Property names used as numeric series on standard charts.

```yaml
Type: String[]
Parameter Sets: Categorical
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Slide
Target slide that will receive the chart (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Title
Optional chart title.

```yaml
Type: String
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Type
Chart type to create.

```yaml
Type: PowerPointChartType
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values: ClusteredColumn, Line, Pie, Doughnut, Scatter

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Width
Chart width in points.

```yaml
Type: Double
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -X
Left offset in points from the slide origin.

```yaml
Type: Double
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -XProperty
Property name used for the X axis on scatter charts.

```yaml
Type: String
Parameter Sets: Scatter
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Y
Top offset in points from the slide origin.

```yaml
Type: Double
Parameter Sets: Default, Categorical, Scatter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -YProperty
Property names used as numeric Y series on scatter charts.

```yaml
Type: String[]
Parameter Sets: Scatter
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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointChart`

## RELATED LINKS

- None
