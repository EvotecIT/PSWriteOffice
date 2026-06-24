---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelPageBreak
## SYNOPSIS
Gets manual row and column page breaks from Excel worksheets.

## SYNTAX
### Context (Default)
```powershell
Get-OfficeExcelPageBreak [-Sheet <string>] [-SheetIndex <int>] [-Row] [-Column] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelPageBreak [-InputPath] <string> [-Sheet <string>] [-SheetIndex <int>] [-Row] [-Column] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelPageBreak -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Row] [-Column] [<CommonParameters>]
```

## DESCRIPTION
Gets manual row and column page breaks from Excel worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $breaks = Get-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data
$breaks |
    Sort-Object Type, Position
```

Returns row and column page-break records for print-layout audits.

## PARAMETERS

### -Column
Only return column page breaks.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to inspect outside the DSL context.

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
Workbook path to inspect.

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

### -Row
Only return row page breaks.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: Nullable`1
Parameter Sets: Context, Path, Document
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
