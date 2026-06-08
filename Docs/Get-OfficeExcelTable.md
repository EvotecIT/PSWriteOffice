---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelTable
## SYNOPSIS
Gets Excel tables defined in a workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelTable [-InputPath] <string> [-Name <string>] [-Sheet <string>] [-SheetIndex <int>] [<CommonParameters>]
```

### Uri
```powershell
Get-OfficeExcelTable [-Uri] <uri> [-AllowHttp] [-Name <string>] [-Sheet <string>] [-SheetIndex <int>] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelTable -Document <ExcelDocument> [-Name <string>] [-Sheet <string>] [-SheetIndex <int>] [<CommonParameters>]
```

## DESCRIPTION
Gets Excel tables defined in a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $tables = Get-OfficeExcelTable -Path .\report.xlsx -Sheet Data
$tables |
    Select-Object -Property Name, Sheet, Range |
    Export-Csv -Path .\ExcelTables.csv -NoTypeInformation
```

Returns table metadata for workbook documentation or generated-artifact proof.

## PARAMETERS

### -AllowHttp
Allow HTTP workbook downloads in addition to HTTPS.

```yaml
Type: SwitchParameter
Parameter Sets: Uri
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
Optional table name filter.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Optional sheet name filter.

```yaml
Type: String
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Optional sheet index (0-based) filter.

```yaml
Type: Nullable`1
Parameter Sets: Path, Uri, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Uri
Remote workbook URI to inspect.

```yaml
Type: Uri
Parameter Sets: Uri
Aliases: Url
Possible values:

Required: True
Position: 0
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
