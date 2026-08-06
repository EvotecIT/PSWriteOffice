---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelColumn
## SYNOPSIS
Writes values or formatting to a column in the current worksheet.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelColumn [[-Column] <Int32>] [-ColumnName <string>] [-Values <Object[]>] [-StartRow <int>] [-Width <Double>] [-Hidden <Boolean>] [-AutoFit] [<CommonParameters>]
```

## DESCRIPTION
Writes values or formatting to a column in the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelColumn -Column 1 -Values 'North','South' -AutoFit }
```

Writes values into column A and adjusts the width.

## PARAMETERS

### -AutoFit
Auto-fit the column width after updates.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
1-based column index.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ColumnName
Column letter reference (e.g., A, BC).

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: ColumnLetter, Letter
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Hide or show the column.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StartRow
Starting row index (1-based) for values.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Values
Values to write down the column.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Column width to apply.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None
