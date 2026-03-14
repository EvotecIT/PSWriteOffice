---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Protect-OfficeExcelSheet
## SYNOPSIS
Protects the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Protect-OfficeExcelSheet [-AllowSelectLockedCells <bool>] [-AllowSelectUnlockedCells <bool>] [-AllowFormatCells <bool>] [-AllowFormatColumns <bool>] [-AllowFormatRows <bool>] [-AllowInsertColumns <bool>] [-AllowInsertRows <bool>] [-AllowInsertHyperlinks <bool>] [-AllowDeleteColumns <bool>] [-AllowDeleteRows <bool>] [-AllowSort <bool>] [-AllowAutoFilter <bool>] [-AllowPivotTables <bool>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Protect-OfficeExcelSheet -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-AllowSelectLockedCells <bool>] [-AllowSelectUnlockedCells <bool>] [-AllowFormatCells <bool>] [-AllowFormatColumns <bool>] [-AllowFormatRows <bool>] [-AllowInsertColumns <bool>] [-AllowInsertRows <bool>] [-AllowInsertHyperlinks <bool>] [-AllowDeleteColumns <bool>] [-AllowDeleteRows <bool>] [-AllowSort <bool>] [-AllowAutoFilter <bool>] [-AllowPivotTables <bool>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Protects the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Protect-OfficeExcelSheet }
```

Enables worksheet protection.

## PARAMETERS

### -AllowAutoFilter
Allow AutoFilter.

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

### -AllowDeleteColumns
Allow deleting columns.

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

### -AllowDeleteRows
Allow deleting rows.

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

### -AllowFormatCells
Allow formatting cells.

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

### -AllowFormatColumns
Allow formatting columns.

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

### -AllowFormatRows
Allow formatting rows.

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

### -AllowInsertColumns
Allow inserting columns.

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

### -AllowInsertHyperlinks
Allow inserting hyperlinks.

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

### -AllowInsertRows
Allow inserting rows.

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

### -AllowPivotTables
Allow PivotTables.

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

### -AllowSelectLockedCells
Allow selecting locked cells.

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

### -AllowSelectUnlockedCells
Allow selecting unlocked cells.

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

### -AllowSort
Allow sorting.

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

### -PassThru
Emit the worksheet after protection.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

