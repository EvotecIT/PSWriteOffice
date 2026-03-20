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
Add-OfficePowerPointChart -Data <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Slide <PowerPointSlide>] [-Type <PowerPointChartType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Title <string>] [<CommonParameters>]
```

### Scatter
```powershell
Add-OfficePowerPointChart -Data <Object[]> -XProperty <string> -YProperty <string[]> [-Slide <PowerPointSlide>] [-Type <PowerPointChartType>] [-X <double>] [-Y <double>] [-Width <double>] [-Height <double>] [-Title <string>] [<CommonParameters>]
```

## DESCRIPTION
Adds a chart to a PowerPoint slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficePowerPointChart -Slide $slide -Data $rows -CategoryProperty Month -SeriesProperty Sales,Profit -Title 'Monthly performance'
```

Creates a clustered column chart using Month for categories and Sales/Profit as series.

### EXAMPLE 2
```powershell
PS>Add-OfficePowerPointChart -Slide $slide -Type Scatter -Data $rows -XProperty Quarter -YProperty Revenue -Title 'Revenue trend'
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

### -Data
Source objects used to build chart data.

```yaml
Type: Object[]
Parameter Sets: Categorical, Scatter
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

