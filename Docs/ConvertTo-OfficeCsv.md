---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeCsv
## SYNOPSIS
Converts objects or a CSV document into CSV text or a file.

## SYNTAX
### InputObject (Default)
```powershell
ConvertTo-OfficeCsv [-InputObject <Object>] [-Delimiter <char>] [-IncludeHeader <bool>] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-OutputPath <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeCsv -Document <CsvDocument> [-Delimiter <char>] [-IncludeHeader <bool>] [-NewLine <string>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-OutputPath <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Converts objects or a CSV document into CSV text or a file.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>$csv = $data | ConvertTo-OfficeCsv
```

Generates CSV text from the input objects.

### EXAMPLE 2
```powershell
PS>$rows = @(
[ordered]@{ Id = 1; Name = 'Alpha'; Total = 10.5 },
[ordered]@{ Id = 2; Name = 'Beta'; Total = 7.25 }
)
$rows | ConvertTo-OfficeCsv -OutputPath .\export.csv -Delimiter ';'
```

Uses ordered dictionaries to enforce column order and a custom delimiter.

### EXAMPLE 3
```powershell
PS>$data | ConvertTo-OfficeCsv -IncludeHeader:$false -OutputPath .\noheader.csv
```

Writes rows only when a downstream system expects headerless CSV.

## PARAMETERS

### -Culture
Culture used for value formatting.

```yaml
Type: CultureInfo
Parameter Sets: InputObject, Document
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
Parameter Sets: InputObject, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
CSV document to serialize.

```yaml
Type: CsvDocument
Parameter Sets: Document
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
Parameter Sets: InputObject, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeHeader
Include the header row in the output.

```yaml
Type: Boolean
Parameter Sets: InputObject, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Objects to convert into CSV rows.

```yaml
Type: Object
Parameter Sets: InputObject
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -NewLine
Override the newline sequence.

```yaml
Type: String
Parameter Sets: InputObject, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Optional output path for the CSV file.

```yaml
Type: String
Parameter Sets: InputObject, Document
Aliases: OutPath
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit a FileInfo when saving to disk.

```yaml
Type: SwitchParameter
Parameter Sets: InputObject, Document
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

- `OfficeIMO.CSV.CsvDocument
System.Object`

## OUTPUTS

- `System.String
System.IO.FileInfo`

## RELATED LINKS

- None

