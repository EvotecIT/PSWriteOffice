---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Import-OfficeCsv
## SYNOPSIS
Imports CSV rows as PSCustomObjects or dictionaries.

## SYNTAX
### PathDelimiter (Default)
```powershell
Import-OfficeCsv [-Path] <string[]> [-NoHeader] [-Header <string[]>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

### Document
```powershell
Import-OfficeCsv [-Document <CsvDocument>] [-AsHashtable] [<CommonParameters>]
```

### PathCulture
```powershell
Import-OfficeCsv [-Path] <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

### PathDetect
```powershell
Import-OfficeCsv [-Path] <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

### LiteralPathDelimiter
```powershell
Import-OfficeCsv -LiteralPath <string[]> [-NoHeader] [-Header <string[]>] [-Delimiter <char>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

### LiteralPathCulture
```powershell
Import-OfficeCsv -LiteralPath <string[]> -UseCulture [-NoHeader] [-Header <string[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

### LiteralPathDetect
```powershell
Import-OfficeCsv -LiteralPath <string[]> -DetectDelimiter [-NoHeader] [-Header <string[]>] [-DelimiterCandidates <char[]>] [-TrimWhitespace <bool>] [-AllowEmptyLines] [-SkipCommentRowsBeforeHeader <bool>] [-SkipCommentRows] [-CommentCharacter <char>] [-RecognizeW3CFieldsHeader <bool>] [-ColumnCountMismatchPolicy <CsvColumnCountMismatchPolicy>] [-Mode <CsvLoadMode>] [-Culture <cultureinfo>] [-Encoding <Encoding>] [-AsHashtable] [<CommonParameters>]
```

## DESCRIPTION
Uses the CSV header to map fields to property names.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Import-OfficeCsv -Path .\data.csv | Format-Table
```

Imports each row as a PSCustomObject.

### EXAMPLE 2
```powershell
PS> Import-OfficeCsv -Path .\data.csv -AsHashtable | ForEach-Object { $_['Name'] }
```

Uses hashtables for dynamic schemas or key-based access.

## PARAMETERS

### -AllowEmptyLines
Allow empty lines in the input.

```yaml
Type: SwitchParameter
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, Document, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, LiteralPathDelimiter
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
Parameter Sets: PathDetect, LiteralPathDetect
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
Parameter Sets: PathDetect, LiteralPathDetect
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
CSV document to read when already loaded.

```yaml
Type: CsvDocument
Parameter Sets: Document
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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

### -Mode
Load mode controlling materialization.

```yaml
Type: CsvLoadMode
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

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
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: True
```

### -RecognizeW3CFieldsHeader
Recognize W3C Extended Log File Format #Fields: rows as headers.

```yaml
Type: Boolean
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
Aliases: None
Possible values:

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
Parameter Sets: PathDelimiter, PathCulture, PathDetect, LiteralPathDelimiter, LiteralPathCulture, LiteralPathDetect
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
Parameter Sets: PathCulture, LiteralPathCulture
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

- `OfficeIMO.CSV.CsvDocument
System.String[]`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None
