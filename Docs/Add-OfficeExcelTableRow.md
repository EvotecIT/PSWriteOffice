---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTableRow
## SYNOPSIS
Appends one or more data rows to an existing Excel table.

## SYNTAX
### Path (Default)
```powershell
Add-OfficeExcelTableRow [-InputPath] <string> [-InputObject] <Object> -TableName <string> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelTableRow [-InputObject] <Object> -Document <ExcelDocument> -TableName <string> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

### Table
```powershell
Add-OfficeExcelTableRow [-InputObject] <Object> -Table <ExcelTable> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Use this command when a workbook already contains a named table and the script should extend that
table without recreating the worksheet through the Excel DSL. The command accepts a workbook path,
an open ExcelDocument, or an existing ExcelTable object. Objects,
dictionaries, DataTable, DataView, IDataReader, and DataRow input are
normalized through the same table input pipeline used by Add-OfficeExcelTable.

When a path is supplied, PSWriteOffice opens the workbook, appends the rows, saves the workbook, and
releases the document. When an open workbook or table is supplied, the caller controls the lifetime
and should close or save the workbook after all edits are complete.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeExcel -Path .\Report.xlsx
$doc | Add-OfficeExcelTableRow -Sheet Data -TableName Sales -InputObject ([pscustomobject]@{ Region='APAC'; Revenue=300 })
$doc | Close-OfficeExcel -Save
```

Uses the existing OfficeIMO Excel table append API and keeps the workbook open for further changes.

### EXAMPLE 2
```powershell
PS> $rows = @(
    [pscustomobject]@{ Service='Identity'; Status='Ready'; Owner='IAM' }
    [pscustomobject]@{ Service='Network'; Status='Investigating'; Owner='Platform' }
)
Add-OfficeExcelTableRow -Path .\Readiness.xlsx -Sheet Readiness -TableName ServiceReadiness -InputObject $rows
```

Opens the workbook from disk, appends both objects to the named table, and saves the file.

## PARAMETERS

### -Document
Open workbook to update. The caller remains responsible for saving and closing it.

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

### -InputObject
Rows to append. Accepts objects, dictionaries, DataTables, DataViews, IDataReaders, and DataRows.

```yaml
Type: Object
Parameter Sets: Path, Document, Table
Aliases: Data, Values
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -InputPath
Workbook path to open, update, save, and close.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated table wrapper so additional table operations can continue in the pipeline.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document, Table
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name that owns the table. Use this when table names might repeat across sheets.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Zero-based worksheet index that owns the table.

```yaml
Type: Nullable`1
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Table
Existing OfficeIMO Excel table wrapper to append to when the table has already been resolved.

```yaml
Type: ExcelTable
Parameter Sets: Table
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -TableName
Existing table name or display name, for example the name returned by Get-OfficeExcelTable.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: Name
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

- `OfficeIMO.Excel.ExcelDocument
OfficeIMO.Excel.ExcelTable
System.Object`

## OUTPUTS

- `OfficeIMO.Excel.ExcelTable`

## RELATED LINKS

- None
