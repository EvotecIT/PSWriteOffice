---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Clear-OfficeExcelRange
## SYNOPSIS
Clears values, formulas, styles, and range metadata from an Excel worksheet range.

## SYNTAX
### Context (Default)
```powershell
Clear-OfficeExcelRange -Range <string> [-Sheet <string>] [-SheetIndex <int>] [-Contents] [-Values] [-Formulas] [-Styles] [-Comments] [-Hyperlinks] [-DataValidations] [-ConditionalFormatting] [-Merges] [-Sparklines] [-All] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Clear-OfficeExcelRange [-InputPath] <string> -Range <string> [-Sheet <string>] [-SheetIndex <int>] [-Contents] [-Values] [-Formulas] [-Styles] [-Comments] [-Hyperlinks] [-DataValidations] [-ConditionalFormatting] [-Merges] [-Sparklines] [-All] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Clear-OfficeExcelRange -Document <ExcelDocument> -Range <string> [-Sheet <string>] [-SheetIndex <int>] [-Contents] [-Values] [-Formulas] [-Styles] [-Comments] [-Hyperlinks] [-DataValidations] [-ConditionalFormatting] [-Merges] [-Sparklines] [-All] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Clears values, formulas, styles, and range metadata from an Excel worksheet range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Clear-OfficeExcelRange -Path .\Report.xlsx -Sheet Staging -Range B2:D20 -Contents -Hyperlinks -Confirm:$false
            Get-OfficeExcelRange -Path .\Report.xlsx -Sheet Staging -Range B2:D20 |
                Select-Object Address, Value, Formula
```

Removes values and formulas from the selected range and saves the workbook.

## PARAMETERS

### -All
Clear all supported cell data and range metadata.

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

### -Comments
Clear comments in the selected range.

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

### -ConditionalFormatting
Clear conditional formatting rules that overlap the selected range.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: ConditionalFormats
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Contents
Clear values and formulas.

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

### -DataValidations
Clear data validation rules that overlap the selected range.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: Validation, Validations
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

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

### -Formulas
Clear formulas.

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

### -Hyperlinks
Clear hyperlinks that overlap the selected range.

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

### -Merges
Clear merged-cell definitions that overlap the selected range.

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
A1 range to clear.

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

### -Sheet
Worksheet name to update. Defaults to the current DSL sheet or the first workbook sheet.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) to update. Defaults to the current DSL sheet or the first workbook sheet.

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

### -Sparklines
Clear sparklines whose target cells overlap the selected range.

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

### -Styles
Clear cell style indexes.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: Formats
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Values
Clear literal cell values.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
