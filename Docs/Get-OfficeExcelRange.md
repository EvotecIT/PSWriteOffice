---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelRange
## SYNOPSIS
Reads an explicit A1 range from an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelRange [-InputPath] <string> -Range <string> [-Sheet <string>] [-SheetIndex <int>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelRange -Document <ExcelDocument> -Range <string> [-Sheet <string>] [-SheetIndex <int>] [-HeadersInFirstRow <bool>] [-NumericAsDecimal] [-AsHashtable] [-AsDataTable] [<CommonParameters>]
```

## DESCRIPTION
Reads a specific rectangular range from a workbook. By default it uses the first row as headers and returns each remaining row as a PSCustomObject.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeExcelRange -Path .\Report.xlsx -Sheet 'Data' -Range 'A1:C10'
```

Returns rows from `A1:C10` as PSCustomObjects.

### EXAMPLE 2
```powershell
PS>Get-OfficeExcelRange -Path .\Report.xlsx -SheetIndex 0 -Range 'A1:C10' -AsDataTable
```

Returns the selected range as a raw `DataTable`.

## PARAMETERS

### -AsDataTable
Emit the raw `DataTable`.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AsHashtable
Emit each row as a hashtable instead of a PSCustomObject.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
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
Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -HeadersInFirstRow
Use the first row in the selected range as column headers.

```yaml
Type: Boolean
Parameter Sets: Path, Document
Aliases: None
Required: False
Position: named
Default value: True
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
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
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Range
Explicit A1 range to read.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name to read.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
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
Parameter Sets: Path, Document
Aliases: None
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
- `System.Collections.Hashtable`
- `System.Data.DataTable`

## RELATED LINKS

- None
