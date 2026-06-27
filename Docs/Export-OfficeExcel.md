---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeExcel
## SYNOPSIS
Exports PowerShell objects to an Excel workbook using an operator-friendly surface.

## SYNTAX
### Create (Default)
```powershell
Export-OfficeExcel [-Path] <string> [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <string>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-NoClobber] [-ExcludeProperty <string[]>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Append
```powershell
Export-OfficeExcel [-Path] <string> -Append [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <string>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-AppendToTable] [-ExcludeProperty <string[]>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### ClearSheet
```powershell
Export-OfficeExcel [-Path] <string> -ClearSheet [-InputObject <Object>] [-WorksheetName <string>] [-TableName <string>] [-TableStyle <string>] [-ShowFirstColumn] [-ShowLastColumn] [-NoRowStripes] [-ShowColumnStripes] [-StartRow <int>] [-StartColumn <int>] [-NoHeader] [-NoTable] [-NoAutoFilter] [-AutoFit] [-FreezeTopRow] [-FreezeFirstColumn] [-BoldTopRow] [-Title <string>] [-ExcludeProperty <string[]>] [-ColumnFormat <hashtable>] [-TextColumn <string[]>] [-NumberColumn <string[]>] [-IntegerColumn <string[]>] [-PercentColumn <string[]>] [-CurrencyColumn <string[]>] [-DateColumn <string[]>] [-DateTimeColumn <string[]>] [-FormatDecimals <int>] [-FormatCultureName <string>] [-IncludeHeaderInColumnFormat] [-AutoFitFormattedColumn] [-IgnoreMissingColumnFormat] [-IncludeUnexportableProperties] [-PropertyConversionErrorAction <ActionPreference>] [-Open] [-PassThru] [-DocumentTitle <string>] [-Author <string>] [-Subject <string>] [-Keywords <string>] [-Description <string>] [-Category <string>] [-Company <string>] [-Manager <string>] [-ApplicationName <string>] [-LastModifiedBy <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Provides a fast PowerShell export path while keeping OfficeIMO as the workbook engine.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Data -AutoFit -FreezeTopRow
```

Creates a workbook, writes the objects as a table, auto-fits columns, and freezes the header row.

### EXAMPLE 2
```powershell
PS> $rows | Export-OfficeExcel -Path .\Report.xlsx -WorksheetName Data -TableName Sales -TextColumn Id -CurrencyColumn Revenue -ColumnFormat @{ Rate = @{ Style = 'Percent'; Decimals = 1 }; Created = 'Date' } -FormatCultureName en-US -AutoFitFormattedColumn
```

Formats ID values as text, Revenue as currency, Rate as a one-decimal percentage, and Created as a short date while keeping formatting logic in OfficeIMO.

## PARAMETERS

### -Append
Append rows to an existing worksheet when the workbook exists.

```yaml
Type: SwitchParameter
Parameter Sets: Append
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AppendToTable
Require append operations to extend an existing Excel table instead of writing after the used range.

```yaml
Type: SwitchParameter
Parameter Sets: Append
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ApplicationName
Workbook application-name metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Author
Workbook author metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AutoFit
Auto-fit exported columns.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: AutoSize
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AutoFitFormattedColumn
Auto-fit only columns that receive export-time column formats.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BoldTopRow
Bold the exported header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Category
Workbook category metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ClearCachedFormulaResults
Remove cached formula results before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ClearSheet
Replace the target worksheet inside an existing workbook.

```yaml
Type: SwitchParameter
Parameter Sets: ClearSheet
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnFormat
Header-to-format map. Values may be preset names such as Text, Currency, Percent, Date, or custom Excel number formats.

```yaml
Type: Hashtable
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Company
Workbook company metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CurrencyColumn
Headers that should be formatted as currency.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DateColumn
Headers that should be formatted as dates.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DateSystem
Workbook date system for Excel date serials.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values: 1900, 1904, NineteenHundred, NineteenFour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DateTimeColumn
Headers that should be formatted as date/time values.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Description
Workbook description metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DisableFastPackageWriter
Disable OfficeIMO fast package writers for this save.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DocumentTitle
Workbook document title metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -EvaluateFormulas
Evaluate supported formulas and write cached values before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExcludeProperty
Exclude specific properties from exported objects.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ForceFullCalculationOnOpen
Request a full workbook recalculation when opened in Excel-compatible applications.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FormatCultureName
Culture used by friendly currency column presets, such as en-US or pl-PL.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FormatDecimals
Decimal places used by friendly number, percent, and currency column presets.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FreezeFirstColumn
Freeze the first exported column.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FreezeTopRow
Freeze the exported header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IgnoreMissingColumnFormat
Continue when a requested export-time column format header is missing.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeHeaderInColumnFormat
Include header cells when applying export-time column formats.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeUnexportableProperties
Include properties that cannot be read by exporting a descriptive placeholder value.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Objects to write. Accepts pipeline input.

```yaml
Type: Object
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -IntegerColumn
Headers that should be formatted as whole numbers.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Keywords
Workbook keyword metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LastModifiedBy
Workbook last-modified-by metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Manager
Workbook manager metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MarkFormulasDirty
Mark formula cells dirty before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoAutoFilter
Disable AutoFilter dropdowns on the created table.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoClobber
Do not overwrite an existing workbook unless appending or clearing a sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Create
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeader
Do not emit a header row.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoRowStripes
Disable alternating row stripes for newly created tables.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoTable
Do not create an Excel table around the exported data.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NumberColumn
Headers that should be formatted as decimal numbers.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Open
Open the workbook after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the saved FileInfo.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Destination workbook path.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PercentColumn
Headers that should be formatted as percentages.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PropertyConversionErrorAction
Controls how unreadable PowerShell properties are handled while projecting export rows.

```yaml
Type: ActionPreference
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values: SilentlyContinue, Stop, Continue, Inquire, Ignore, Suspend, Break

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SafePreflight
Run OfficeIMO worksheet preflight cleanup before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SafeRepairDefinedNames
Repair common defined-name issues before saving.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ShowColumnStripes
Enable alternating column stripes for newly created tables.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
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
Parameter Sets: Create, Append, ClearSheet
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
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StartColumn
Starting column for new exports.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StartRow
Starting row for new exports. When appending and left at 1, rows are written after the used range.

```yaml
Type: Int32
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Subject
Workbook subject metadata.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableName
Optional Excel table name.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TableStyle
Built-in Excel table style name.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TextColumn
Headers that should be formatted as text, useful for IDs, zip codes, and leading-zero values.

```yaml
Type: String[]
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Title
Write a title above the exported table.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ValidateOpenXml
Validate the saved package with OpenXmlValidator and throw on errors.

```yaml
Type: SwitchParameter
Parameter Sets: Create, Append, ClearSheet
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WorksheetName
Worksheet name to create or update.

```yaml
Type: String
Parameter Sets: Create, Append, ClearSheet
Aliases: Sheet
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

- `System.Object`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
