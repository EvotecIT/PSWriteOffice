---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelConditionalIconSet
## SYNOPSIS
Adds an icon set conditional format to a range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelConditionalIconSet [-Range] <string> [-IconSet <string>] [-ShowValue <bool>] [-Reverse <bool>] [-PercentThresholds <double[]>] [-NumberThresholds <double[]>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelConditionalIconSet [-Range] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-IconSet <string>] [-ShowValue <bool>] [-Reverse <bool>] [-PercentThresholds <double[]>] [-NumberThresholds <double[]>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an icon set conditional format to a range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelConditionalIconSet -Range 'E2:E50' -IconSet ThreeTrafficLights1 }
```

Applies a traffic-light icon set.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

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

### -IconSet
Icon set to apply.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NumberThresholds
Number thresholds matching the icon count.

```yaml
Type: Double[]
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
Emit the range after applying the format.

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

### -PercentThresholds
Percent thresholds (0..100) matching the icon count.

```yaml
Type: Double[]
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Range
A1 range to format.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Reverse
Reverse the icon order.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Sheet
Worksheet name when using Document.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Nullable`1
Parameter Sets: Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowValue
Show the underlying values.

```yaml
Type: Boolean
Parameter Sets: Context, Document
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

