---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeExcelDelimitedText
## SYNOPSIS
Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.

## SYNTAX
### Path (Default)
```powershell
Import-OfficeExcelDelimitedText [-InputPath] <string> -SourcePath <string> [-Delimiter <char>] [-SheetName <string>] [-CultureName <string>] [-NoHeader] [-NoTable] [-NoTypeConversion] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Import-OfficeExcelDelimitedText -Document <ExcelDocument> -SourcePath <string> [-Delimiter <char>] [-SheetName <string>] [-CultureName <string>] [-NoHeader] [-NoTable] [-NoTypeConversion] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Imports normalized CSV/TSV text into an Excel workbook through OfficeIMO.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $result = Import-OfficeExcelDelimitedText -Path .\Report.xlsx `
                -SourcePath .\sales-pl.csv `
                -Delimiter ';' `
                -CultureName 'pl-PL' `
                -SheetName Sales `
                -PassThru
            $result | Format-List SheetName,Range,RowCount,ColumnCount,Delimiter
```

Normalizes delimited text through OfficeIMO, performs culture-aware value conversion, and writes the result as an Excel table unless -NoTable is used.

## PARAMETERS

### -CultureName
Culture name for number and date conversion.

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

### -Delimiter
Delimiter character. When omitted, it is detected.

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

### -Document
Workbook document.

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
Workbook path.

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

### -NoHeader
Treat the first row as data.

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

### -NoTable
Import rows without creating an Excel table.

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

### -NoTypeConversion
Keep imported values as text.

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

### -PassThru
Emit a result object.

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

### -SheetName
Worksheet name to create or inspect.

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

### -SourcePath
Delimited text source path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
