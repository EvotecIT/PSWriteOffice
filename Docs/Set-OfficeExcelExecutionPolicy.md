---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelExecutionPolicy
## SYNOPSIS
Configures OfficeIMO Excel execution and validation behavior for a workbook.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelExecutionPolicy [-Mode <string>] [-ParallelThreshold <int>] [-MaxDegreeOfParallelism <int>] [-WorksheetValidation <string>] [-Diagnostics] [-DisableAutoFitImmediateSave] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelExecutionPolicy [-Document] <ExcelDocument> [-Mode <string>] [-ParallelThreshold <int>] [-MaxDegreeOfParallelism <int>] [-WorksheetValidation <string>] [-Diagnostics] [-DisableAutoFitImmediateSave] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures OfficeIMO Excel execution and validation behavior for a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $workbook | Set-OfficeExcelExecutionPolicy -Mode Sequential
```

Disables automatic parallel execution decisions for subsequent OfficeIMO operations.

## PARAMETERS

### -Diagnostics
Request diagnostics-aware validation without wiring verbose callbacks.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DisableAutoFitImmediateSave
Do not save worksheet parts immediately after AutoFit width/height mutations.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook whose execution policy should be updated.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -MaxDegreeOfParallelism
Optional cap for parallel compute phases.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Mode
Execution mode for large operations.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: Automatic, Sequential, Parallel

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ParallelThreshold
Global item threshold above which Automatic mode switches to Parallel.

```yaml
Type: Nullable`1
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the workbook after updating the policy.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WorksheetValidation
Worksheet validation mode for write operations.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: Disabled, DiagnosticsOnly, DebugOnly, Always

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

- `OfficeIMO.Excel.ExcelDocument`

## RELATED LINKS

- None
