---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelPowerQueryMetadata
## SYNOPSIS
Adds safe Power Query/connection metadata for Excel-compatible applications to own and refresh.

## SYNTAX
### Path (Default)
```powershell
Add-OfficeExcelPowerQueryMetadata [-InputPath] <string> -Name <string> [-WorksheetName <string>] [-QueryTableName <string>] [-Description <string>] [-CommandText <string>] [-RefreshOnOpen] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPowerQueryMetadata -Document <ExcelDocument> -Name <string> [-WorksheetName <string>] [-QueryTableName <string>] [-Description <string>] [-CommandText <string>] [-RefreshOnOpen] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds safe Power Query/connection metadata for Excel-compatible applications to own and refresh.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeExcelPowerQueryMetadata -Path .\Report.xlsx `
                -Name SalesQuery `
                -WorksheetName Data `
                -CommandText 'let Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content] in Source' `
                -Description 'Sales query metadata authored by automation' `
                -RefreshOnOpen `
                -PassThru
```

Writes package metadata only. OfficeIMO does not execute Power Query M; Excel-compatible applications perform refresh when opened.

## PARAMETERS

### -CommandText
Power Query M expression stored as metadata.

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

### -Description
Connection description.

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
Workbook document to update.

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

### -Name
Connection name stored in workbook metadata.

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

### -PassThru
Emit metadata about the authored package parts.

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

### -QueryTableName
Optional query-table name.

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

### -RefreshOnOpen
Request refresh-on-open metadata.

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

### -WorksheetName
Worksheet that should own query-table metadata. Defaults to the current DSL sheet when used inside New-OfficeExcel.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: Sheet, SheetName, Worksheet
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
