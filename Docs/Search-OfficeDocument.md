---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Search-OfficeDocument
## SYNOPSIS
Searches one Reader result or every supported document below file and folder paths.

## SYNTAX
### Document (Default)
```powershell
Search-OfficeDocument [-InputObject] <OfficeDocumentReadResult> [-Query] <string> [-MatchCase] [-WholeWord] [-MaximumResults <int>] [-AllResults] [<CommonParameters>]
```

### Path
```powershell
Search-OfficeDocument [-Path] <string[]> [-Query] <string> [-MatchCase] [-WholeWord] [-MaximumResults <int>] [-AllResults] [-Recurse] [-Extension <string[]>] [-MaxDocuments <int>] [-NoDocumentLimit] [-MaxStoreItems <int>] [-AllStoreItems] [-MaxDegreeOfParallelism <int>] [-IncludePageLocations] [-StopOnError] [-Reader <OfficeDocumentReader>] [<CommonParameters>]
```

## DESCRIPTION
Searches one Reader result or every supported document below file and folder paths.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Search-OfficeDocument -Path .\Evidence -Recurse -Query 'retention period'
```

Automatically reads supported Word, Excel, PowerPoint, PDF, email, PST, OST, and other registered formats.

### EXAMPLE 2
```powershell
PS> Search-OfficeDocument -Path .\Evidence -Recurse -Query 'invoice' -NoDocumentLimit -AllStoreItems -AllResults
```

Unlimited modes are explicit because very large stores and document collections can consume substantial resources.

## PARAMETERS

### -AllResults
Return every occurrence from each document instead of applying the default result ceiling.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -AllStoreItems
Project every matching item from each email store.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Extension
Optional extensions to include. Registered Reader formats are used automatically when omitted.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: Extensions
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludePageLocations
Compute Word and RTF page locations when supported.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputObject
Normalized document returned by Get-OfficeDocument.

```yaml
Type: OfficeDocumentReadResult
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -MatchCase
Use case-sensitive ordinal matching.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxDegreeOfParallelism
Maximum document reads in flight.

```yaml
Type: Nullable`1
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxDocuments
Maximum documents accepted in one search. The default is 500.

```yaml
Type: Nullable`1
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaximumResults
Maximum occurrences returned per document. The default is 1,000.

```yaml
Type: Int32
Parameter Sets: Document, Path
Aliases: MaxResultsPerDocument
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaxStoreItems
Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.

```yaml
Type: Nullable`1
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoDocumentLimit
Remove the document-count ceiling.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
File, directory, or wildcard path to search.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: FullName, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: True
```

### -Query
Text to find in normalized document blocks.

```yaml
Type: String
Parameter Sets: Document, Path
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Reader
Advanced immutable Reader configured by a .NET host or New-OfficeDocumentReader.

```yaml
Type: OfficeDocumentReader
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Recurse
Search subdirectories.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -StopOnError
Terminate the search when one document cannot be read. The default reports the error and continues.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WholeWord
Return only occurrences surrounded by non-word characters.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path
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

- `OfficeIMO.Reader.OfficeDocumentReadResult
System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentSearchResult
PSWriteOffice.Models.Reader.OfficeDocumentSearchMatch` — PowerShell-friendly occurrence returned by a path-based document search.

## RELATED LINKS

- None
