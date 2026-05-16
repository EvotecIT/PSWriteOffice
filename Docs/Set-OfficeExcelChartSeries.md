---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartSeries
## SYNOPSIS
Configures Excel chart series colors, line style, and markers.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartSeries -Chart <ExcelChart> -SeriesIndex <int> [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <double>] [-MarkerStyle <string>] [-MarkerSize <int>] [-MarkerFillColor <string>] [-MarkerLineColor <string>] [-MarkerLineWidthPoints <double>] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartSeries -Chart <ExcelChart> -SeriesName <string> [-IgnoreCase <bool>] [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <double>] [-MarkerStyle <string>] [-MarkerSize <int>] [-MarkerFillColor <string>] [-MarkerLineColor <string>] [-MarkerLineWidthPoints <double>] [<CommonParameters>]
```

## DESCRIPTION
Configures Excel chart series colors, line style, and markers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartSeries -SeriesName 'Revenue' -FillColor '#4472C4' -LineColor '#1F4E79' -MarkerStyle Circle -MarkerSize 6
```

Applies fill, line, and marker settings to the Revenue series.

## PARAMETERS

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

### -FillColor
Series fill color in hex format.

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

### -LineColor
Series line color in hex format.

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
Series line width in points.

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

### -MarkerFillColor
Marker fill color in hex format.

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

### -MarkerLineColor
Marker line color in hex format.

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

### -MarkerLineWidthPoints
Marker line width in points.

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

### -MarkerSize
Marker size.

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

### -MarkerStyle
Marker style name.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None

