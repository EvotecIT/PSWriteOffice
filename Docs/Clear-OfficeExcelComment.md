---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Clear-OfficeExcelComment
## SYNOPSIS
Clears legacy worksheet comments (notes) that match a filter.

## SYNTAX
### Context (Default)
```powershell
Clear-OfficeExcelComment [-Sheet <string>] [-SheetIndex <int>] [-Address <string>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Clear-OfficeExcelComment [-InputPath] <string> [-Sheet <string>] [-SheetIndex <int>] [-Address <string>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Clear-OfficeExcelComment -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Address <string>] [-Range <string>] [-Author <string>] [-TextContains <string>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Clears legacy worksheet comments (notes) that match a filter.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $removed = Clear-OfficeExcelComment -Path .\Report.xlsx -Sheet Data -TextContains review -Confirm:$false -PassThru
Get-OfficeExcelCommentAudit -Path .\Report.xlsx -IncludeComments |
    Select-Object LegacyCommentCount, ThreadedCommentCount
```

Removes matching comments and saves the workbook.

## PARAMETERS

### -Address
A1 cell address to match.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -All
Allow clearing all comments on the selected worksheet(s) when no filter is supplied.

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

### -Author
Comment author to match, ignoring case.

```yaml
Type: String
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
Workbook to update outside the DSL context.

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
Workbook path to update.

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
Returns the number of comments cleared.

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

### -Range
A1 cell or range to match.

```yaml
Type: String
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
Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.

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

### -TextContains
Text fragment to match, ignoring case.

```yaml
Type: String
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

- `System.Int32`

## RELATED LINKS

- None
