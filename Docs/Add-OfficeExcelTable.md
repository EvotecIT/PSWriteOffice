---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTable
## SYNOPSIS
Writes tabular data to the current worksheet and formats it as an Excel table.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelTable [-InputObject] <Object> [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-View <OfficeTableView>] [-TableName <string>] [-TableStyle <string>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-NoAutoFilter] [-AutoFit] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
ExcelSheet 'Data' { Add-OfficeExcelTable -InputObject $data -TableName 'Sales' }
```

Writes two rows and formats them as a styled Excel table.

## PARAMETERS

### -AutoFit
Auto-fit the table columns after insertion.

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

### -InputObject
Source objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences to convert into table rows.

```yaml
Type: Object
Parameter Sets: __AllParameterSets
Aliases: Data, DataTable
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -NoAutoFilter
Disable AutoFilter dropdowns.

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

### -NoHeader
Skip writing headers.

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

### -NoRowStripes
Disable alternating row stripes for the created table.

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

### -PassThru
Return the created range string.

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

### -ShowColumnStripes
Enable alternating column stripes for the created table.

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

### -ShowFirstColumn
Emphasize the first table column when the selected style supports it.

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

### -ShowLastColumn
Emphasize the last table column when the selected style supports it.

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

### -StartColumn
Starting column for the data (1-based).

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

### -StartRow
Starting row for the data (1-based).

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

### -TableName
Name to assign to the table.

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

### -TableStyle
Built-in table style to apply.

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

### -View
Projection to apply before writing the table.

```yaml
Type: OfficeTableView
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Normal, Transpose

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
