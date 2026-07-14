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
### PathDelimiter (Default)
```powershell
Get-OfficeCsv [-Path] <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### PathCulture
```powershell
Get-OfficeCsv [-Path] <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### PathDetect
```powershell
Get-OfficeCsv [-Path] <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### LiteralPathDelimiter
```powershell
Get-OfficeCsv -LiteralPath <string[]> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### LiteralPathCulture
```powershell
Get-OfficeCsv -LiteralPath <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### LiteralPathDetect
```powershell
Get-OfficeCsv -LiteralPath <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-CompressionType <CsvCompressionType>] [-MaxDecompressedBytes <long>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### TextDelimiter
```powershell
Get-OfficeCsv -Text <string> [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-Delimiter <char>] [-DelimiterText <string>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### TextCulture
```powershell
Get-OfficeCsv -Text <string> -UseCulture [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

### TextDetect
```powershell
Get-OfficeCsv -Text <string> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-SkipRows <int>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-DuplicateHeaderBehavior <CsvDuplicateHeaderBehavior>] [-NullValue <string>] [-DateTimeFormats <string[]>] [-QuoteParsingMode <CsvQuoteParsingMode>] [-StaticColumns <IDictionary>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-ParseErrorAction <CsvParseErrorAction>] [-CollectParseErrors] [-MaxParseErrors <int>] [-MaxFieldLength <int>] [-MaxQuotedFieldLength <int>] [-NormalizeQuotes] [-InternStrings] [-ProgressInterval <int>] [<CommonParameters>]
```

## DESCRIPTION
Returns an CsvDocument for inspection or further transformations.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $csv = Get-OfficeCsv -Path .\data.csv
```

Loads the CSV file into an OfficeIMO CsvDocument.

### EXAMPLE 2
```powershell
PS> $csv = Get-OfficeCsv -Text \"Name;Total`nAlpha;10\" -Delimiter ';'
```

Parses a semicolon-delimited CSV string into a document.

### EXAMPLE 3
```powershell
PS> $csv = Get-OfficeCsv -Path .\data.csv; $csv.Header
```

Returns the header list so you can verify the expected column names.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CollectParseErrors
Collect parse errors and write them as non-terminating errors after each input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -CompressionType
Compression used when reading files. Auto infers from the file extension.

```yaml
Type: CsvCompressionType
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values: None, Auto, GZip, Deflate, Brotli, ZLib

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DateTimeFormats
Additional date/time formats used by typed conversions and validation.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter, TextDelimiter
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
Parameter Sets: PathDetect, LiteralPathDetect, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DelimiterText
Field delimiter text for multi-character delimiters such as || or ::.

```yaml
Type: String
Parameter Sets: PathDelimiter, LiteralPathDelimiter, TextDelimiter
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
Parameter Sets: PathDetect, LiteralPathDetect, TextDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -DuplicateHeaderBehavior
Controls how duplicate header names are handled.

```yaml
Type: CsvDuplicateHeaderBehavior
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Preserve, Rename, Throw

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Header
Explicit header names to use; when provided, the first CSV record is treated as data.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InternStrings
Reuse repeated string values through a per-read cache.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -LiteralPath
Literal path to one or more CSV files.

```yaml
Type: String[]
Parameter Sets: LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: PSPath, LP
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: True
```

### -MaxDecompressedBytes
Maximum decompressed bytes to read from compressed CSV files.

```yaml
Type: Nullable`1
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxFieldLength
Maximum length allowed for any parsed field.

```yaml
Type: Nullable`1
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxParseErrors
Maximum number of collected parse errors before parsing fails.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxQuotedFieldLength
Maximum length allowed for fields parsed from quoted records.

```yaml
Type: Nullable`1
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NormalizeQuotes
Normalize curly quote characters to straight quotes.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NullValue
Token that is materialized as null when loading rows.

```yaml
Type: String
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ParseErrorAction
How parse errors are handled.

```yaml
Type: CsvParseErrorAction
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Throw, SkipRow

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to one or more CSV files. Wildcards are supported.

```yaml
Type: String[]
Parameter Sets: PathDelimiter, PathCulture, PathDetect
Aliases: FilePath, InputPath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: True
```

### -ProgressInterval
Report progress every N parsed records.

```yaml
Type: Nullable`1
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -QuoteParsingMode
Controls whether malformed quoted fields are parsed leniently or rejected.

```yaml
Type: CsvQuoteParsingMode
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values: Lenient, Strict

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SkipCommentRows
Skip comment rows throughout the file.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SkipRows
Number of parsed CSV records to skip before header discovery or data output.

```yaml
Type: Int32
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StaticColumns
Static columns appended to every loaded row.

```yaml
Type: IDictionary
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -TrimWhitespace
Trim whitespace around unquoted fields.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect, TextDelimiter, TextCulture, TextDetect
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
Parameter Sets: PathCulture, LiteralPathCulture, TextCulture
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

- `System.String[]`

## OUTPUTS

- `OfficeIMO.CSV.CsvDocument`

## RELATED LINKS

- None
