---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeCsv
## SYNOPSIS
Exports objects or a CSV document to a CSV file.

## SYNTAX
### InputObjectPathDelimiter (Default)
```powershell
Export-OfficeCsv [-Path] <string> [-InputObject <Object>] [-Delimiter <char>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectPathCulture
```powershell
Export-OfficeCsv [-Path] <string> -UseCulture [-InputObject <Object>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectLiteralPathDelimiter
```powershell
Export-OfficeCsv -LiteralPath <string> [-InputObject <Object>] [-Delimiter <char>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### InputObjectLiteralPathCulture
```powershell
Export-OfficeCsv -LiteralPath <string> -UseCulture [-InputObject <Object>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentPathDelimiter
```powershell
Export-OfficeCsv [-Path] <string> -Document <CsvDocument> [-Delimiter <char>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentPathCulture
```powershell
Export-OfficeCsv [-Path] <string> -Document <CsvDocument> -UseCulture [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentLiteralPathDelimiter
```powershell
Export-OfficeCsv -Document <CsvDocument> -LiteralPath <string> [-Delimiter <char>] [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentLiteralPathCulture
```powershell
Export-OfficeCsv -Document <CsvDocument> -LiteralPath <string> -UseCulture [-NoHeader] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-FormulaInjectionPolicy <CsvFormulaInjectionPolicy>] [-UseQuotes <CsvQuoteMode>] [-QuoteFields <string[]>] [-PassThru] [-Append] [-NoClobber] [-Force] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Use ConvertTo-OfficeCsv when the destination should be CSV text in the pipeline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $data | Export-OfficeCsv -Path .\export.csv
```

Streams PowerShell objects into a CSV file.

### EXAMPLE 2
```powershell
PS> $data | Export-OfficeCsv -Path .\export.csv -UseCulture -Culture pl-PL
```

Uses the selected culture list separator as the delimiter.

## PARAMETERS

### -Append
Append rows to an existing CSV file. Existing headers are reused when present.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Culture
Culture used for value formatting.

```yaml
Type: CultureInfo
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Delimiter
Field delimiter character.

```yaml
Type: Char
Parameter Sets: InputObjectPathDelimiter, InputObjectLiteralPathDelimiter, DocumentPathDelimiter, DocumentLiteralPathDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
CSV document to export.

```yaml
Type: CsvDocument
Parameter Sets: DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Encoding
Encoding used when writing files.

```yaml
Type: Encoding
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Force
Allow overwriting read-only files and appending rows with missing existing columns.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FormulaInjectionPolicy
Controls how formula-like values are written.

```yaml
Type: CsvFormulaInjectionPolicy
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: Preserve, Escape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Objects to export into CSV rows.

```yaml
Type: Object
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -LiteralPath
Literal destination CSV path. Wildcards are not expanded.

```yaml
Type: String
Parameter Sets: InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: PSPath, LP
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NewLine
Override the newline sequence.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoClobber
Do not overwrite an existing CSV file.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: NoOverwrite
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeader
Omit the header row from the output.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo for the exported file.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Destination CSV path.

```yaml
Type: String
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, DocumentPathDelimiter, DocumentPathCulture
Aliases: FilePath, OutputPath, OutPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -QuoteFields
Field names that should always be quoted when UseQuotes is AsNeeded.

```yaml
Type: String[]
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -UseCulture
Use the list separator from the selected or current culture as the delimiter.

```yaml
Type: SwitchParameter
Parameter Sets: InputObjectPathCulture, InputObjectLiteralPathCulture, DocumentPathCulture, DocumentLiteralPathCulture
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -UseQuotes
Controls when CSV fields are quoted. Defaults to quoting only fields that need it.

```yaml
Type: CsvQuoteMode
Parameter Sets: InputObjectPathDelimiter, InputObjectPathCulture, InputObjectLiteralPathDelimiter, InputObjectLiteralPathCulture, DocumentPathDelimiter, DocumentPathCulture, DocumentLiteralPathDelimiter, DocumentLiteralPathCulture
Aliases: None
Possible values: AsNeeded, Always, Never

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object
OfficeIMO.CSV.CsvDocument`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
