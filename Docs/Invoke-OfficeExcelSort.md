---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelSort
## SYNOPSIS
Sorts the used range on the current worksheet.

## SYNTAX
### ContextSingle (Default)
```powershell
Invoke-OfficeExcelSort [-Header] <string> [-Descending] [-PassThru] [<CommonParameters>]
```

### DocumentSingle
```powershell
Invoke-OfficeExcelSort [-Header] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <int>] [-Descending] [-PassThru] [<CommonParameters>]
```

### DocumentOrder
```powershell
Invoke-OfficeExcelSort -Document <ExcelDocument> -Order <hashtable> [-Sheet <string>] [-SheetIndex <int>] [-PassThru] [<CommonParameters>]
```

### ContextOrder
```powershell
Invoke-OfficeExcelSort -Order <hashtable> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Sorts the used range on the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>ExcelSheet 'Data' { Invoke-OfficeExcelSort -Header 'Name' }
```

Sorts by the Name column in ascending order.

### EXAMPLE 2
```powershell
PS>$order = [ordered]@{ Status = $true; Total = $false }\nExcelSheet 'Data' { Invoke-OfficeExcelSort -Order $order }
```

Sorts by Status ascending, then Total descending.

## PARAMETERS

### -Descending
Sort descending (single-column sort).

```yaml
Type: SwitchParameter
Parameter Sets: ContextSingle, DocumentSingle
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
Parameter Sets: DocumentSingle, DocumentOrder
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Header
Header to sort by (single-column sort).

```yaml
Type: String
Parameter Sets: ContextSingle, DocumentSingle
Aliases: None
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Order
Ordered dictionary of header => ascending (true/false).

```yaml
Type: Hashtable
Parameter Sets: DocumentOrder, ContextOrder
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the worksheet after sorting.

```yaml
Type: SwitchParameter
Parameter Sets: ContextSingle, DocumentSingle, DocumentOrder, ContextOrder
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
Parameter Sets: DocumentSingle, DocumentOrder
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
Parameter Sets: DocumentSingle, DocumentOrder
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

