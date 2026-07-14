---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentVisual
## SYNOPSIS
Reads visual payloads discovered by OfficeIMO.Reader from a supported document.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentVisual [-Path] <string> [-AsExport] [-OutputDirectory <string>] [-NoOverwrite] [-Indented] [-NoPayload] [-NoJson] [-MaxInputBytes <long>] [-OpenXmlMaxCharactersInPart <long>] [-MaxChars <int>] [-MaxTableRows <int>] [-ExcludeWordFootnotes] [-ExcludePowerPointNotes] [-NoExcelHeaders] [-ExcelChunkRows <int>] [-ExcelSheetName <string>] [-ExcelA1Range <string>] [-NoMarkdownHeadingChunks] [-NoHashes] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Reads visual payloads discovered by OfficeIMO.Reader from a supported document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $visuals = Get-OfficeDocumentVisual -Path .\Deck.pptx -AsExport -OutputDirectory .\reader-visuals -Indented
$visuals | Select-Object Id, ContentType, PayloadPath
```

Extracts supported visual payloads and writes payload plus metadata sidecars.

## PARAMETERS

### -AsExport
Return deterministic payload and JSON export bundles instead of visual models.

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

### -Indented
Indent JSON payloads in export bundles and sidecars.

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
Maximum table rows per emitted table.

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

### -NoJson
Do not write JSON sidecars when -OutputDirectory is used.

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

### -NoOverwrite
Do not overwrite existing sidecar files.

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

### -NoPayload
Do not write raw payload sidecars when -OutputDirectory is used.

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

### -OutputDirectory
Optional directory where visual sidecars should be written.

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

### -Path
File path to read.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
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

- `OfficeIMO.Reader.ReaderVisual
OfficeIMO.Reader.ReaderVisualExportBundle
OfficeIMO.Reader.ReaderVisualMaterializedExport`

## RELATED LINKS

- None
