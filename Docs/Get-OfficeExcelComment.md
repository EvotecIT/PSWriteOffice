---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelComment
## SYNOPSIS
Gets legacy worksheet comments (notes) from one or more worksheets.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelComment [-InputPath] <string> [-Sheet <string>] [-SheetIndex <int>] [-Address <string>] [-Range <string>] [-Author <string>] [-TextContains <string>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelComment -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Address <string>] [-Range <string>] [-Author <string>] [-TextContains <string>] [<CommonParameters>]
```

## DESCRIPTION
Gets legacy worksheet comments (notes) from one or more worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $comments = Get-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -TextContains review
$comments |
    Select-Object SheetName, Address, Author, Text
```

Returns matching comment metadata without modifying the workbook.

## PARAMETERS

### -Address
A1 cell address to match.

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

### -Author
Comment author to match, ignoring case.

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

### -Range
A1 cell or range to match.

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

### -Sheet
Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.

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

### -SheetIndex
Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.

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

### -TextContains
Text fragment to match, ignoring case.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
