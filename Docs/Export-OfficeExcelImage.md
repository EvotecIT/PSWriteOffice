---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeExcelImage
## SYNOPSIS
Exports workbook sheets as PNG or SVG images with one result per sheet.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeExcelImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-Options <ExcelWorkbookImageExportOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeExcelImage [-OutputPath] <string> -Document <ExcelDocument> [-Format <OfficeImageExportFormat>] [-Options <ExcelWorkbookImageExportOptions>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports workbook sheets as PNG or SVG images with one result per sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeExcelImage -Path .\Report.xlsx -OutputPath .\Images
```

Writes one image per selected sheet and returns OfficeImageExportResult objects.

## PARAMETERS

### -Document
Open workbook instance.

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

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, Document
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Options
Optional sheet selection, range, size, and rendering settings.

```yaml
Type: ExcelWorkbookImageExportOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Destination folder.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to the workbook.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
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

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None
