---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Join-OfficeExcelWorkbook
## SYNOPSIS
Imports selected or all worksheets from one Excel workbook into another.

## SYNTAX
### Path (Default)
```powershell
Join-OfficeExcelWorkbook [-InputPath] <string> [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-SourceSheet <string[]>] [-SheetNamePrefix <string>] [<CommonParameters>]
```

### Document
```powershell
Join-OfficeExcelWorkbook -Document <ExcelDocument> [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-SourceSheet <string[]>] [-SheetNamePrefix <string>] [<CommonParameters>]
```

## DESCRIPTION
Imports selected or all worksheets from one Excel workbook into another.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $merge = Join-OfficeExcelWorkbook -Path .\Target.xlsx -SourcePath .\Source.xlsx -SourceSheet Data,Summary -SheetNamePrefix 'Imported '
Get-OfficeExcelSummary -Path .\Target.xlsx |
    Select-Object Path, WorksheetCount
```

Copies worksheets from Source.xlsx into Target.xlsx using OfficeIMO workbook merge logic.

## PARAMETERS

### -Document
Target workbook to update outside the DSL context.

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

### -InputPath
Target workbook path to update.

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

### -SheetNamePrefix
Prefix added to every imported worksheet name.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourceDocument
Optional source workbook object.

```yaml
Type: ExcelDocument
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourcePath
Optional source workbook path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourceSheet
Specific source worksheet names to import. Defaults to all source sheets.

```yaml
Type: String[]
Parameter Sets: Path, Document
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

- `OfficeIMO.Excel.ExcelWorkbookMergeResult`

## RELATED LINKS

- None
