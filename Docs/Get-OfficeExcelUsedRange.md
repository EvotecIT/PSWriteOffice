---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelUsedRange
## SYNOPSIS
Reads the used range from an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelUsedRange [-InputPath] <string> [-Sheet <string>] [-SheetIndex <int>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelUsedRange -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

## DESCRIPTION
Reads the used range from an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeExcelUsedRange -Path .\report.xlsx -Sheet 'Data'
```

Reads the sheet's used range, treating the first row as headers.

## PARAMETERS

### -AsDataTable
Emit the raw DataTable instead of row objects.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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
Workbook to inspect.

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

### -HeadersInFirstRow
Use the first row as column headers.

```yaml
Type: Boolean
Parameter Sets: Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NumericAsDecimal
Prefer decimals instead of doubles for numeric values.

```yaml
Type: SwitchParameter
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
Worksheet name to read; defaults to the first sheet.

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
Zero-based worksheet index to read; defaults to the first sheet.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject
System.Collections.Hashtable
System.Data.DataTable`

## RELATED LINKS

- None

