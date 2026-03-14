---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelNamedRange
## SYNOPSIS
Gets defined names (named ranges) from an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelNamedRange [-InputPath] <string> [-Name <string>] [-Sheet <string>] [-SheetIndex <int>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelNamedRange -Document <ExcelDocument> [-Name <string>] [-Sheet <string>] [-SheetIndex <int>] [<CommonParameters>]
```

## DESCRIPTION
Gets defined names (named ranges) from an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeExcelNamedRange -Path .\report.xlsx
```

Returns workbook-level named ranges.

## PARAMETERS

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

### -Name
Optional named range to retrieve.

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
Optional sheet name for sheet-scoped names.

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
Optional sheet index (0-based) for sheet-scoped names.

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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None

