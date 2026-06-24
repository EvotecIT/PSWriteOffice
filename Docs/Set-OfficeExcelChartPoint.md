---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartPoint
## SYNOPSIS
Configures fill and line styling for a single Excel chart data point.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartPoint -Chart <ExcelChart> -SeriesIndex <int> -PointIndex <uint> [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <double>] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartPoint -Chart <ExcelChart> -SeriesName <string> -PointIndex <uint> [-IgnoreCase <bool>] [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <double>] [<CommonParameters>]
```

## DESCRIPTION
Configures fill and line styling for a single Excel chart data point.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartPoint -SeriesName 'Revenue' -PointIndex 1 -FillColor '#C00000'
```

Applies a point-specific fill override to the second point in the Revenue series.

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
Point fill color in hex format.

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
Point line color in hex format.

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
Point line width in points.

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

### -PointIndex
Zero-based data point index within the series.

```yaml
Type: UInt32
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: True
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
