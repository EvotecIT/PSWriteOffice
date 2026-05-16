---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Copy-OfficeExcelSheet
## SYNOPSIS
Copies a worksheet within a workbook or from another workbook.

## SYNTAX
### Context (Default)
```powershell
Copy-OfficeExcelSheet [[-SourceSheet] <string>] [-NewName] <string> [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-ValidationMode <SheetNameValidationMode>] [<CommonParameters>]
```

### Path
```powershell
Copy-OfficeExcelSheet [-InputPath] <string> [[-SourceSheet] <string>] [-NewName] <string> [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-ValidationMode <SheetNameValidationMode>] [<CommonParameters>]
```

### Document
```powershell
Copy-OfficeExcelSheet [[-SourceSheet] <string>] [-NewName] <string> -Document <ExcelDocument> [-SourceDocument <ExcelDocument>] [-SourcePath <string>] [-ValidationMode <SheetNameValidationMode>] [<CommonParameters>]
```

## DESCRIPTION
Copies a worksheet within a workbook or from another workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Copy-OfficeExcelSheet -Path .\Report.xlsx -SourceSheet Data -NewName DataCopy
```

Creates a copy of the Data worksheet.

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

### -NewName
Name for the copied worksheet.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: Name, DestinationSheet
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourceDocument
Optional source workbook object for cross-workbook copies.

```yaml
Type: ExcelDocument
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SourcePath
Optional source workbook path for cross-workbook copies.

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

### -SourceSheet
Worksheet to copy. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: Sheet, WorksheetName
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ValidationMode
Controls how invalid destination sheet names are handled.

```yaml
Type: SheetNameValidationMode
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: None, Sanitize, Strict

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

- `OfficeIMO.Excel.ExcelSheet`

## RELATED LINKS

- None
