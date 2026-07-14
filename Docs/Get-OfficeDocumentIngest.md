---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentIngest
## SYNOPSIS
Reads a folder into an OfficeIMO.Reader ingestion summary.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentIngest [-FolderPath] <string> [-NoRecurse] [-MaxFiles <int>] [-MaxTotalBytes <long>] [-Extension <string[]>] [-NoChunks] [-MaxInputBytes <long>] [-OpenXmlMaxCharactersInPart <long>] [-MaxChars <int>] [-MaxTableRows <int>] [-ExcludeWordFootnotes] [-ExcludePowerPointNotes] [-NoExcelHeaders] [-ExcelChunkRows <int>] [-ExcelSheetName <string>] [-ExcelA1Range <string>] [-NoMarkdownHeadingChunks] [-NoHashes] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Reads a folder into an OfficeIMO.Reader ingestion summary.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $ingest = Get-OfficeDocumentIngest -FolderPath .\Reports -Extension docx,pdf,rtf -MaxFiles 50
$ingest.Files | Select-Object Path, Status, ChunkCount
```

Reads supported files from a folder and returns the ingestion summary with per-file status and chunk counts.

## PARAMETERS

### -ExcelA1Range
Optional Excel A1 range to read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExcelChunkRows
Excel rows per emitted worksheet chunk.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExcelSheetName
Optional Excel sheet name to read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ExcludePowerPointNotes
Exclude PowerPoint speaker notes.

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

### -ExcludeWordFootnotes
Exclude Word footnotes.

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

### -Extension
Allowed folder extensions such as .docx, .xlsx, .pdf, .html, or json.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: Extensions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -FolderPath
Folder path to ingest.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -MaxChars
Maximum emitted chunk characters.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxFiles
Maximum number of folder files to read.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxInputBytes
Maximum input size in bytes.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxTableRows
Maximum table rows per emitted table chunk.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxTotalBytes
Maximum total folder bytes to read.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoChunks
Do not materialize chunks in the returned ingestion result.

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

### -NoExcelHeaders
Treat the first Excel row as data instead of headers.

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

### -NoHashes
Disable source and chunk hash computation.

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

### -NoMarkdownHeadingChunks
Do not split Markdown by headings.

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

### -NoRecurse
Do not recurse into child folders.

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

### -OpenXmlMaxCharactersInPart
OpenXML maximum characters per part.

```yaml
Type: Nullable`1
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Reader
{{ Fill Reader Description }}

```yaml
Type: OfficeDocumentReader
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

- `System.String`

## OUTPUTS

- `OfficeIMO.Reader.ReaderIngestResult`

## RELATED LINKS

- None
