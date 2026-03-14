---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelComment
## SYNOPSIS
Adds or replaces a comment (note) on a worksheet cell.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelComment -Text <string> [-Row <int>] [-Column <int>] [-Address <string>] [-Author <string>] [-Initials <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelComment -Document <ExcelDocument> -Text <string> [-Sheet <string>] [-SheetIndex <int>] [-Row <int>] [-Column <int>] [-Address <string>] [-Author <string>] [-Initials <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or replaces a comment (note) on a worksheet cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Add-OfficeExcelComment -Address 'A1' -Text 'Review this cell' }
```

Creates a cell comment at A1.

## PARAMETERS

### -Address
A1-style cell address (e.g., A1, C5).

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

### -Author
Author name (optional).

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

### -Column
1-based column index.

```yaml
Type: Nullable`1
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

### -Initials
Author initials (optional).

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

### -PassThru
Emit the worksheet after adding the comment.

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

### -Row
1-based row index.

```yaml
Type: Nullable`1
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

### -Text
Comment text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values: 

Required: True
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

