---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelRow
## SYNOPSIS
Writes a row of values to the current worksheet.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeExcelRow [-Row] <int> [-Values] <Object[]> [-StartColumn <int>] [<CommonParameters>]
```

## DESCRIPTION
Writes a row of values to the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Set-OfficeExcelRow -Row 2 -Values 'North', 1200 }
```

Writes two values into row 2, columns A and B.

## PARAMETERS

### -Row
1-based row index.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StartColumn
Starting column index (1-based).

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Values
Values to write across the row.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

