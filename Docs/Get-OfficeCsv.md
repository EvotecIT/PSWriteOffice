---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeCsv
## SYNOPSIS
Loads a CSV document from disk or parses CSV text.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeCsv [-InputPath] <string> [-HasHeaderRow <bool>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [<CommonParameters>]
```

### Text
```powershell
Get-OfficeCsv -Text <string> [-HasHeaderRow <bool>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [<CommonParameters>]
```

## DESCRIPTION
Loads a CSV document from disk or parses CSV text.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$csv = Get-OfficeCsv -Path .\data.csv
```

Loads the CSV file into an OfficeIMO CsvDocument.

### EXAMPLE 2
```powershell
PS>$csv = Get-OfficeCsv -Text \"Name;Total`nAlpha;10\" -Delimiter ';'
```

Parses a semicolon-delimited CSV string into a document.

### EXAMPLE 3
```powershell
PS>$csv = Get-OfficeCsv -Path .\data.csv; $csv.Header
```

Returns the header list so you can verify the expected column names.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Text
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Culture
Culture used for type conversions.

```yaml
Type: CultureInfo
Parameter Sets: Path, Text
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
Parameter Sets: Path, Text
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Encoding
Encoding used when reading the file.

```yaml
Type: Encoding
Parameter Sets: Path, Text
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -HasHeaderRow
Indicates whether the first record is a header row.

```yaml
Type: Boolean
Parameter Sets: Path, Text
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the CSV file.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Mode
Load mode controlling materialization.

```yaml
Type: CsvLoadMode
Parameter Sets: Path, Text
Aliases: None
Possible values: InMemory, Stream

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
CSV text to parse.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: Path, Text
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

- `OfficeIMO.CSV.CsvDocument`

## RELATED LINKS

- None

