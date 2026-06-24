---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelPackageMetadata
## SYNOPSIS
Adds explicit workbook package metadata such as connection or query-table XML.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelPackageMetadata -Kind <string> -Xml <string> [-WorksheetName <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelPackageMetadata [-InputPath] <string> -Kind <string> -Xml <string> [-WorksheetName <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPackageMetadata -Document <ExcelDocument> -Kind <string> -Xml <string> [-WorksheetName <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Adds explicit workbook package metadata such as connection or query-table XML.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeExcelPackageMetadata -Path .\Report.xlsx -Kind Connection -Xml $connectionsXml
            Add-OfficeExcelPackageMetadata -Path .\Report.xlsx -Kind QueryTable -WorksheetName Data -Xml $queryTableXml
```

Adds caller-supplied XML metadata parts. OfficeIMO preserves these parts but does not refresh external queries.

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

### -Kind
Metadata kind to add.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: Connection, QueryTable

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit metadata about the added package part.

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

### -WorksheetName
Worksheet for query-table metadata. Defaults to the current DSL sheet, or the first worksheet outside the DSL.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: Sheet, SheetName, Worksheet
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Xml
XML payload to add as package metadata.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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
