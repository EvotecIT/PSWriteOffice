---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartStyle
## SYNOPSIS
Applies a built-in style and color preset to an Excel chart.

## SYNTAX
```powershell
Set-OfficeExcelChartStyle [-Chart] <ExcelChart> [-StyleId <int>] [-ColorStyleId <int>] [<CommonParameters>]
```

## DESCRIPTION
Applies one of the built-in OfficeIMO chart style and color presets to an existing chart.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$chart | Set-OfficeExcelChartStyle
```

Applies the default OfficeIMO chart style preset.

### EXAMPLE 2
```powershell
PS>$chart | Set-OfficeExcelChartStyle -StyleId 251 -ColorStyleId 10
```

Applies an explicit style and color preset combination.

## PARAMETERS

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: (All)
Aliases: None
Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -StyleId
Chart style identifier.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: 251
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColorStyleId
Chart color style identifier.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases: None
Required: False
Position: named
Default value: 10
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- [Set-OfficeExcelChartLegend](Set-OfficeExcelChartLegend.md)
- [Set-OfficeExcelChartDataLabels](Set-OfficeExcelChartDataLabels.md)
