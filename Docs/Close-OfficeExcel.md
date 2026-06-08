---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Close-OfficeExcel
## SYNOPSIS
Closes an Excel workbook and optionally saves it.

## SYNTAX
### __AllParameterSets
```powershell
Close-OfficeExcel -Document <ExcelDocument> [-Save] [-Path <string>] [-Show] [-Password <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [<CommonParameters>]
```

## DESCRIPTION
Convenience wrapper so scripts do not need to call Save or Dispose directly.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeExcel -Path .\report.xlsx {
    Add-OfficeExcelSheet -Name Data {
        Set-OfficeExcelRow -Row 1 -Values 'Region', 'Revenue'
        Set-OfficeExcelRow -Row 2 -Values 'EMEA', 98000
    }
}
$workbook = Get-OfficeExcel -Path .\report.xlsx
$workbook | Close-OfficeExcel -Save -Path .\report-final.xlsx -SafePreflight -ValidateOpenXml
```

Saves pending changes through OfficeIMO's normal save path, validates the package, and releases the workbook.

## PARAMETERS

### -Document
Workbook to close.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Password
Password used to save the workbook as an encrypted package.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Optional output path when saving.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SafePreflight
Run OfficeIMO worksheet preflight cleanup before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SafeRepairDefinedNames
Repair common defined-name issues before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Save
Persist changes before closing.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Show
Open the workbook in Excel after saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ValidateOpenXml
Validate the saved package with OpenXmlValidator and throw on errors.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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
