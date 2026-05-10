---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordChart
## SYNOPSIS
Adds a chart to a Word document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeWordChart -Data <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeWordChart -Data <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Document <WordDocument>] [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Paragraph
```powershell
Add-OfficeWordChart -Data <Object[]> -CategoryProperty <string> -SeriesProperty <string[]> [-Paragraph <WordParagraph>] [-Type <WordChartType>] [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Creates a Word chart from object data using one category property and one or more numeric series properties.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordChart -Type Pie -Data $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Revenue mix'
```

Creates a pie chart using Region labels and Revenue as the slice values.

### EXAMPLE 2
```powershell
PS> Add-OfficeWordChart -Document $doc -Type Line -Data $rows -CategoryProperty Month -SeriesProperty Sales,Profit -Legend
```

Creates a multi-series line chart on the document and shows a legend.

## PARAMETERS

### -CategoryProperty
Property name used for category labels.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
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
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Target document that will receive the chart.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -FitToPageWidth
Scale the chart width to the page content width.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HeightPixels
Chart height in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Legend
Add a legend to the chart.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LegendPosition
Legend position when -Legend is used.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: Left, Right, Top, Bottom, TopRight

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Paragraph
Target paragraph used as the chart anchor.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -PassThru
Emit the created chart.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SeriesColor
Color values applied to the series in order.

```yaml
Type: String[]
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SeriesProperty
Property names used as numeric series.

```yaml
Type: String[]
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Optional chart title.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
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
Type: WordChartType
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: Pie, Bar, Line, Area

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthFraction
Fraction of the page content width to use when -FitToPageWidth is specified.

```yaml
Type: Double
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WidthPixels
Chart width in pixels.

```yaml
Type: Int32
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -XAxisTitle
Optional X axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -YAxisTitle
Optional Y axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: Context, Document, Paragraph
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

- `OfficeIMO.Word.WordDocument
OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordChart`

## RELATED LINKS

- None

