---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertFrom-OfficeCsv
## SYNOPSIS
Converts CSV text to PSCustomObjects or dictionaries.

## SYNTAX
### TextDelimiter (Default)
```powershell
ConvertFrom-OfficeCsv [-Text] <string> [-NoHeader] [-Header <string[]>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

### TextCulture
```powershell
ConvertFrom-OfficeCsv [-Text] <string> -UseCulture [-NoHeader] [-Header <string[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

### TextDetect
```powershell
ConvertFrom-OfficeCsv [-Text] <string> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-AsHashtable] [<CommonParameters>]
```

## DESCRIPTION
Reads CSV text from -Text or the pipeline and maps rows by header.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $rows = ConvertFrom-OfficeCsv -Text "Name,Value`nAlpha,1"
```

Parses CSV text and emits row objects without writing a temporary file.

### EXAMPLE 2
```powershell
PS> "Name,Value", "Alpha,1" | ConvertFrom-OfficeCsv
```

Treats piped lines as one CSV stream.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnCountMismatchPolicy
Controls how rows with fewer or more fields than the header are handled.

```yaml
Type: CsvColumnCountMismatchPolicy
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Strict, PadMissingFieldsAndIgnoreExtraFields

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CommentCharacter
Character that identifies comment rows.

```yaml
Type: Char
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextDelimiter
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DelimiterCandidates
Delimiter candidates to consider when detecting the delimiter.

```yaml
Type: Char[]
Parameter Sets: TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DetectDelimiter
Detect the delimiter from the first meaningful records.

```yaml
Type: SwitchParameter
Parameter Sets: TextDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Header
Explicit header names to use; when provided, the first CSV record is treated as data.

```yaml
Type: String[]
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Mode
Load mode controlling materialization.

```yaml
Type: CsvLoadMode
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: InMemory, Stream

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHeader
Treat the first record as data and generate default column names.

```yaml
Type: SwitchParameter
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -RecognizeW3CFieldsHeader
Recognize W3C Extended Log File Format #Fields: rows as headers.

```yaml
Type: Boolean
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SkipCommentRows
Skip comment rows throughout the input.

```yaml
Type: SwitchParameter
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SkipCommentRowsBeforeHeader
Skip comment rows starting with # while discovering the header.

```yaml
Type: Boolean
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

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
Parameter Sets: TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: TextCulture
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

- `System.String`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
