---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelWriteReservation
## SYNOPSIS
Sets workbook write-reservation metadata.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelWriteReservation [-ReadOnlyRecommended] [-UserName <string>] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelWriteReservation [-Path] <string> [-ReadOnlyRecommended] [-UserName <string>] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelWriteReservation -Document <ExcelDocument> [-ReadOnlyRecommended] [-UserName <string>] [-Password <string>] [-LegacyPasswordHash <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets workbook write-reservation metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Set-OfficeExcelWriteReservation -Path .\Report.xlsx -ReadOnlyRecommended -UserName 'Reporting Team' -PassThru |
                Format-List ReadOnlyRecommended, UserName
```

Writes Excel file-sharing/write-reservation metadata without encrypting the file or protecting workbook structure.

## PARAMETERS

### -Document
Open workbook document to update.

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

### -LegacyPasswordHash
Optional precomputed legacy write-reservation hash.

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

### -PassThru
Emit the updated write-reservation metadata.

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
Optional write-reservation password. This is legacy Excel metadata, not package encryption.

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

### -Path
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, InputPath, FullName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ReadOnlyRecommended
Recommend opening the workbook as read-only.

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

### -UserName
User name stored in the write-reservation metadata.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
