---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelTemplate
## SYNOPSIS
Applies Excel template markers such as {{Name}} to one or more worksheets.

## SYNTAX
### Context (Default)
```powershell
Invoke-OfficeExcelTemplate -Value <hashtable> [-Sheet <string>] [-SheetIndex <int>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Invoke-OfficeExcelTemplate [-InputPath] <string> -Value <hashtable> [-Sheet <string>] [-SheetIndex <int>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeExcelTemplate -Document <ExcelDocument> -Value <hashtable> [-Sheet <string>] [-SheetIndex <int>] [-CultureName <string>] [-MissingValueBehavior <ExcelTemplateMissingValueBehavior>] [-ThrowOnMissing] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Applies Excel template markers such as {{Name}} to one or more worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Invoke-OfficeExcelTemplate -Path .\Invoice.xlsx -Sheet Invoice -Value @{ Number = 'INV-001'; Total = 123.45 } -CultureName en-US
```

Replaces matching template markers and saves the workbook.

## PARAMETERS

### -CultureName
Culture name used for built-in marker format aliases such as currency and date.

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

### -Document
Workbook to update outside the DSL context.

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

### -InputPath
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MissingValueBehavior
Behavior used when a marker is not supplied by -Value.

```yaml
Type: Nullable`1
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
Returns the number of marker replacements.

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

### -Sheet
Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.

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

### -SheetIndex
Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: Nullable`1
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ThrowOnMissing
Throws when a marker is not supplied by -Value.

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

### -Value
Template marker values keyed by marker name.

```yaml
Type: Hashtable
Parameter Sets: Context, Path, Document
Aliases: Values
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

- `System.Int32`

## RELATED LINKS

- None
