---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelChart
## SYNOPSIS
Adds a chart to the current worksheet using a range or table.

## SYNTAX
### ContextRange (Default)
```powershell
Add-OfficeExcelChart [-Range] <string> -Row <int> -Column <int> [-WidthPixels <int>] [-HeightPixels <int>] [-Type <ExcelChartType>] [-Title <string>] [-HasHeaders <bool>] [-IncludeCachedData <bool>] [-PassThru] [<CommonParameters>]
```

### DocumentRange
```powershell
Add-OfficeExcelChart [-Range] <string> -Document <ExcelDocument> -Row <int> -Column <int> [-Sheet <string>] [-SheetIndex <int>] [-WidthPixels <int>] [-HeightPixels <int>] [-Type <ExcelChartType>] [-Title <string>] [-HasHeaders <bool>] [-IncludeCachedData <bool>] [-PassThru] [<CommonParameters>]
```

### DocumentTable
```powershell
Add-OfficeExcelChart [-TableName] <string> -Document <ExcelDocument> -Row <int> -Column <int> [-Sheet <string>] [-SheetIndex <int>] [-WidthPixels <int>] [-HeightPixels <int>] [-Type <ExcelChartType>] [-Title <string>] [-IncludeCachedData <bool>] [-PassThru] [<CommonParameters>]
```

### ContextTable
```powershell
Add-OfficeExcelChart [-TableName] <string> -Row <int> -Column <int> [-WidthPixels <int>] [-HeightPixels <int>] [-Type <ExcelChartType>] [-Title <string>] [-IncludeCachedData <bool>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a chart to the current worksheet using a range or table.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelChart -Range 'A1:D10' -Row 2 -Column 6 -Type Line -Title 'Trend' }
```

Creates a line chart from A1:D10 and places it at F2.

### EXAMPLE 2
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelChart -TableName 'Sales' -Row 2 -Column 6 -Type ColumnClustered }
```

Creates a chart from the Sales table.

## PARAMETERS

### -Column
Top-left column (1-based) where the chart should be placed.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -HasHeaders
Whether the range includes headers.

```yaml
Type: Boolean
Parameter Sets: ContextRange, DocumentRange
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
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeCachedData
Include cached data in the chart for portability.

```yaml
Type: Boolean
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the created chart.

```yaml
Type: SwitchParameter
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Range
A1 range containing chart data.

```yaml
Type: String
Parameter Sets: ContextRange, DocumentRange
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Row
Top-left row (1-based) where the chart should be placed.

```yaml
Type: Int32
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Nullable`1
Parameter Sets: DocumentRange, DocumentTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableName
Table name containing chart data.

```yaml
Type: String
Parameter Sets: DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Type
Chart type.

```yaml
Type: ExcelChartType
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
Aliases: None
Possible values: ColumnClustered, ColumnStacked, BarClustered, BarStacked, Line, Area, Pie, Doughnut, Scatter, Bubble

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
Parameter Sets: ContextRange, DocumentRange, DocumentTable, ContextTable
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

