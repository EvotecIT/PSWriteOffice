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
### Context
```powershell
Add-OfficeWordChart [-Type <WordChartType>] [-Data] <Object[]> [-CategoryProperty] <string> [-SeriesProperty] <string[]> [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeWordChart -Document <WordDocument> [-Type <WordChartType>] [-Data] <Object[]> [-CategoryProperty] <string> [-SeriesProperty] <string[]> [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

### Paragraph
```powershell
Add-OfficeWordChart -Paragraph <WordParagraph> [-Type <WordChartType>] [-Data] <Object[]> [-CategoryProperty] <string> [-SeriesProperty] <string[]> [-WidthPixels <int>] [-HeightPixels <int>] [-Title <string>] [-SeriesColor <string[]>] [-Legend] [-LegendPosition <string>] [-XAxisTitle <string>] [-YAxisTitle <string>] [-FitToPageWidth] [-WidthFraction <double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a chart to a Word document.
Use `-Paragraph` when you want to anchor the chart in a specific place, including a paragraph created inside a table cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordChart -Type Pie -Data $rows -CategoryProperty Region -SeriesProperty Revenue -Title 'Revenue mix'
```

Creates a pie chart using Region labels and Revenue as the slice values.

### EXAMPLE 2
```powershell
PS>Add-OfficeWordChart -Document $doc -Type Line -Data $rows -CategoryProperty Month -SeriesProperty Sales,Profit -Legend
```

Creates a multi-series line chart on the document and shows a legend.

### EXAMPLE 3
```powershell
PS>$table = Add-OfficeWordTable -InputObject $rows -PassThru
PS>$paragraph = $table.Rows[1].Cells[1].AddParagraph()
PS>Add-OfficeWordChart -Paragraph $paragraph -Type Pie -Data $chartRows -CategoryProperty Region -SeriesProperty Revenue
```

Creates a chart anchored to a paragraph inside a table cell.

## PARAMETERS

### -CategoryProperty
Property name used for category labels.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Data
Source objects used to build chart data.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases: None

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Target document that will receive the chart.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FitToPageWidth
Scale the chart width to the page content width.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeightPixels
Chart height in pixels.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: 360
Accept pipeline input: False
Accept wildcard characters: False
```

### -Legend
Add a legend to the chart.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LegendPosition
Legend position when -Legend is used.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: Right
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Target paragraph used as the chart anchor.

```yaml
Type: WordParagraph
Parameter Sets: Paragraph
Aliases: None

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created chart.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeriesColor
Color values applied to the series in order.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SeriesProperty
Property names used as numeric series.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases: None

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional chart title.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Chart type to create.

```yaml
Type: WordChartType
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: Pie
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthFraction
Fraction of the page content width to use when -FitToPageWidth is specified.

```yaml
Type: Double
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -WidthPixels
Chart width in pixels.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: 600
Accept pipeline input: False
Accept wildcard characters: False
```

### -XAxisTitle
Optional X axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -YAxisTitle
Optional Y axis title for non-pie charts.

```yaml
Type: String
Parameter Sets: (All)
Aliases: None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`
- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordChart`

## RELATED LINKS

- [Add-OfficeWordParagraph](Add-OfficeWordParagraph.md)
