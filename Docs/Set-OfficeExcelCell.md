---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelCell
## SYNOPSIS
Sets a cell value, formula, or number format within the current worksheet.

## SYNTAX
### Coordinates
```powershell
Set-OfficeExcelCell [-Row <int>] [-Column <int>] [-Value <Object>] [-Formula <string>] [-NumberFormat <string>] [<CommonParameters>]
```

### Address
```powershell
Set-OfficeExcelCell [-Address <string>] [-Value <Object>] [-Formula <string>] [-NumberFormat <string>] [<CommonParameters>]
```

## DESCRIPTION
Sets a cell value, formula, or number format within the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Set-OfficeExcelCell -Address 'A1' -Value 'Region'; Set-OfficeExcelCell -Row 1 -Column 2 -Value 'Revenue' }
```

Writes two headers in the first row.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

```yaml
Type: String
Parameter Sets: Address
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Column
1-based column index.

```yaml
Type: Nullable`1
Parameter Sets: Coordinates
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Formula
Formula text (without leading =).

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NumberFormat
Number format code to apply.

```yaml
Type: String
Parameter Sets: Coordinates, Address
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Row
1-based row index.

```yaml
Type: Nullable`1
Parameter Sets: Coordinates
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Value
Cell value to assign.

```yaml
Type: Object
Parameter Sets: Coordinates, Address
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

- `None`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

