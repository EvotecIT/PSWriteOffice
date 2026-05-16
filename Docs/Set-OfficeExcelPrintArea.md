---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelPrintArea
## SYNOPSIS
Sets the print area for a worksheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelPrintArea [-Range] <string> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelPrintArea [-InputPath] <string> [-Range] <string> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelPrintArea [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sets the print area for a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficeExcelPrintArea -Path .\Report.xlsx -Sheet Data -Range A1:H100
```

Stores the worksheet-local Excel print area definition.

## PARAMETERS

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
Emit the worksheet after setting the print area.

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
A1 range to print.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

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
Worksheet index when using a workbook object or path.

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

- `System.Object`

## RELATED LINKS

- None
