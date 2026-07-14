---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeExcel
## SYNOPSIS
Imports rows from an Excel workbook as PowerShell objects.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeExcel [-Path] <string> [-WorksheetName <string>] [-SheetIndex <int>] [-AllSheets] [-Range <string>] [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-NoHeader] [-NumericAsDecimal] [-FormulaMode <string>] [-CultureName <string>] [-AsHashtable] [-AsDataTable] [-AsDataReader] [-ByColumn] [-SchemaSampleSize <int>] [-ChunkRows <int>] [<CommonParameters>]
```

### Uri
```powershell
Import-OfficeExcel [-Uri] <uri> [-AllowHttp] [-WorksheetName <string>] [-SheetIndex <int>] [-AllSheets] [-Range <string>] [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-NoHeader] [-NumericAsDecimal] [-FormulaMode <string>] [-CultureName <string>] [-AsHashtable] [-AsDataTable] [-AsDataReader] [-ByColumn] [-SchemaSampleSize <int>] [-ChunkRows <int>] [<CommonParameters>]
```

### Document
```powershell
Import-OfficeExcel -Document <ExcelDocument> [-WorksheetName <string>] [-SheetIndex <int>] [-AllSheets] [-Range <string>] [-StartRow <int>] [-EndRow <int>] [-StartColumn <int>] [-EndColumn <int>] [-NoHeader] [-NumericAsDecimal] [-FormulaMode <string>] [-CultureName <string>] [-AsHashtable] [-AsDataTable] [-AsDataReader] [-ByColumn] [-SchemaSampleSize <int>] [-ChunkRows <int>] [<CommonParameters>]
```

## DESCRIPTION
Provides a fast PowerShell read command over the OfficeIMO reader pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = Import-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -NumericAsDecimal
$rows |
    Where-Object Status -eq 'Pending' |
    Export-Csv -Path .\PendingRows.csv -NoTypeInformation
```

Reads the used range on the Data worksheet, emits PSCustomObjects, and filters them in PowerShell.

### EXAMPLE 2
```powershell
PS> $rows = Import-OfficeExcel -Path .\Workbook.xlsx -AllSheets
$rows | Group-Object WorksheetName
```

Reads the used range from each worksheet and adds a WorksheetName property to each emitted row.

### EXAMPLE 3
```powershell
PS> Import-OfficeExcel -Path .\Workbook.xlsx -WorksheetName Metrics -ByColumn |
                Where-Object ColumnName -eq 'Revenue' |
                Select-Object -ExpandProperty Values
```

Returns one object per column with the column name, 1-based column index, and the column values as an array.

## PARAMETERS

### -AllowHttp
Allow HTTP workbook downloads in addition to HTTPS.

```yaml
Type: SwitchParameter
Parameter Sets: Uri
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllSheets
Import all worksheets. Each emitted row or column includes WorksheetName.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AsDataReader
Emit a forward-only IDataReader for database bulk-copy workflows.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AsDataTable
Emit a DataTable instead of enumerating row objects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AsHashtable
Emit rows as hashtables instead of PSCustomObjects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ByColumn
Emit one object per column with ColumnName, ColumnIndex, and Values instead of row objects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ChunkRows
Worksheet row count requested from each streaming chunk when -AsDataReader is used.

```yaml
Type: Int32
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CultureName
Culture used when parsing numbers and dates stored as text.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook document to import from.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -EndColumn
Ending column for an explicit rectangular range.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -EndRow
Ending row for an explicit rectangular range.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FormulaMode
Formula read mode. CachedValue returns workbook cached results; FormulaText returns formula expressions when present.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values: CachedValue, FormulaText

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeader
Treat all rows as data and generate column names instead of using the first row as headers.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Workbook path to import.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, InputPath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### -Range
Optional A1 range to read. When omitted, the used range is imported.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### -SchemaSampleSize
Maximum row count inspected when -AsDataReader infers the reader schema.

```yaml
Type: Int32
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Zero-based worksheet index to read.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### -StartColumn
Starting column for an explicit rectangular range.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StartRow
Starting row for an explicit rectangular range.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Uri
Remote workbook URI to import.

```yaml
Type: Uri
Parameter Sets: Uri
Aliases: Url
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### -WorksheetName
Worksheet name to read; defaults to the first sheet.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: Sheet
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String
System.Uri
OfficeIMO.Excel.ExcelDocument
System.Nullable`1[[System.Int32, System.Private.CoreLib, Version=10.0.0.0, Culture=neutral, PublicKeyToken=7cec85d7bea7798e]]`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
