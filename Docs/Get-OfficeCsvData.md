---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeCsvData
## SYNOPSIS
Reads CSV rows as PSCustomObjects or dictionaries.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeCsvData [[-InputPath] <string>] [-Document <CsvDocument>] [-HasHeaderRow <bool>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

## DESCRIPTION
Reads CSV rows as PSCustomObjects or dictionaries.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficeCsvData -Path .\data.csv | Format-Table
```

Returns each row as a PSCustomObject.

### EXAMPLE 2
```powershell
PS>Get-OfficeCsvData -Path .\data.csv -AsHashtable | ForEach-Object { $_['Name'] }
```

Uses hashtables for dynamic schemas or key-based access.

### EXAMPLE 3
```powershell
PS>Get-OfficeCsvData -Path .\data.csv -Delimiter ';' -HasHeaderRow:$false
```

Reads CSV files that lack a header row and use a custom delimiter.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

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

### -AsHashtable
Emit dictionaries instead of PSCustomObjects.

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

### -Culture
Culture used for type conversions.

```yaml
Type: CultureInfo
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
CSV document to read when already loaded.

```yaml
Type: CsvDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Encoding
Encoding used when reading the file.

```yaml
Type: Encoding
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to a CSV file.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, Path
Possible values: 

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Mode
Load mode controlling materialization.

```yaml
Type: CsvLoadMode
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: InMemory, Stream

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.CSV.CsvDocument`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

