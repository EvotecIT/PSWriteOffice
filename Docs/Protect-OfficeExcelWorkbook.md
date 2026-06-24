---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Protect-OfficeExcelWorkbook
## SYNOPSIS
Protects workbook structure or windows metadata. This is not file encryption.

## SYNTAX
### Context (Default)
```powershell
Protect-OfficeExcelWorkbook [-NoStructure] [-ProtectWindows] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
Protect-OfficeExcelWorkbook [-InputPath] <string> [-NoStructure] [-ProtectWindows] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Protect-OfficeExcelWorkbook -Document <ExcelDocument> [-NoStructure] [-ProtectWindows] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Protects workbook structure or windows metadata. This is not file encryption.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Protect-OfficeExcelWorkbook -Path .\Report.xlsx -Password secret
            Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation |
                Select-Object Passed, ProtectionSummary
```

Writes workbook-level structure protection metadata and saves the workbook.

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

### -LegacyPasswordHash
Optional precomputed legacy workbook protection hash to write as-is.

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

### -NoStructure
Do not protect workbook structure.

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

### -PassThru
Emit the workbook after protection.

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

### -Password
Optional workbook protection password. This is UI protection, not package encryption.

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

### -ProtectWindows
Protect workbook windows where supported by the consuming application.

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
