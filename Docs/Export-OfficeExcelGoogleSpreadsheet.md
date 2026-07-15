---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeExcelGoogleSpreadsheet
## SYNOPSIS
Plans, compiles, or exports an Excel workbook to Google Sheets.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeExcelGoogleSpreadsheet [-Path] <string> [-Options <GoogleSheetsSaveOptions>] [-Session <GoogleWorkspaceSession>] [-PlanOnly] [-AsBatch] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeExcelGoogleSpreadsheet -Document <ExcelDocument> [-Options <GoogleSheetsSaveOptions>] [-Session <GoogleWorkspaceSession>] [-PlanOnly] [-AsBatch] [-FailOnLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Plans, compiles, or exports an Excel workbook to Google Sheets.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeExcelGoogleSpreadsheet -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Export-OfficeExcelGoogleSpreadsheet -Document 'Value'
```


## PARAMETERS

### -AsBatch
Return the provider-neutral request batch without contacting Google.

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
Excel workbook to export.

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

### -FailOnLoss
Throw when translation reports a warning or error.

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

### -Options
Google Sheets translation and destination settings.

```yaml
Type: GoogleSheetsSaveOptions
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to an Excel workbook.

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

### -PlanOnly
Return the translation plan without compiling requests or contacting Google.

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

### -Session
Configured Google Workspace session used for a live export.

```yaml
Type: GoogleWorkspaceSession
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

- `OfficeIMO.Excel.GoogleSheets.GoogleSheetsTranslationPlan
OfficeIMO.Excel.GoogleSheets.GoogleSheetsBatch
OfficeIMO.Excel.GoogleSheets.GoogleSpreadsheetReference`

## RELATED LINKS

- None
