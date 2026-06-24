---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelPrintLayout
## SYNOPSIS
Applies a reusable worksheet print layout preset.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelPrintLayout [-Sheet <string>] [-SheetIndex <int>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <ExcelPageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <uint>] [-FitToHeight <uint>] [-Scale <uint>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <int>] [-RepeatLastRow <int>] [-RepeatFirstColumn <int>] [-RepeatLastColumn <int>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelPrintLayout [-InputPath] <string> [-Sheet <string>] [-SheetIndex <int>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <ExcelPageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <uint>] [-FitToHeight <uint>] [-Scale <uint>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <int>] [-RepeatLastRow <int>] [-RepeatFirstColumn <int>] [-RepeatLastColumn <int>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelPrintLayout -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Preset <ExcelPrintLayoutPreset>] [-PrintArea <string>] [-Orientation <ExcelPageOrientation>] [-Margins <ExcelMarginPreset>] [-FitToWidth <uint>] [-FitToHeight <uint>] [-Scale <uint>] [-PageOrder <ExcelPageOrder>] [-RepeatFirstRow <int>] [-RepeatLastRow <int>] [-RepeatFirstColumn <int>] [-RepeatLastColumn <int>] [-NoPresetPrintTitles] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies a reusable worksheet print layout preset.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Report' { Set-OfficeExcelPrintLayout -Preset Report -PrintArea A1:H40 }
```

Applies landscape orientation, narrow margins, one-page-wide scaling, and repeated header row.

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

### -FitToHeight
Optional pages-tall fit override. Use 0 for unlimited height.

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

### -FitToWidth
Optional pages-wide fit override.

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

### -Margins
Optional margin preset override.

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

### -NoPresetPrintTitles
Do not apply print-title rows from the selected preset.

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

### -Orientation
Optional orientation override.

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

### -PageOrder
Optional multi-page print order override.

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

### -PassThru
Emit the worksheet after applying the layout.

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

### -Preset
Print layout preset.

```yaml
Type: ExcelPrintLayoutPreset
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: Worksheet, Report, Dashboard, DataTable

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PrintArea
Optional print area in A1 notation.

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

### -RepeatFirstColumn
Optional first 1-based repeated print-title column.

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

### -RepeatFirstRow
Optional first 1-based repeated print-title row.

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

### -RepeatLastColumn
Optional last 1-based repeated print-title column.

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

### -RepeatLastRow
Optional last 1-based repeated print-title row.

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

### -Scale
Optional manual scale percentage override.

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

### -Sheet
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index when using a workbook object or path.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelSheet
System.Management.Automation.PSObject`

## RELATED LINKS

- None
