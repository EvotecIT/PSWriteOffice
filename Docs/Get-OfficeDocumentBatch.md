---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentBatch
## SYNOPSIS
Reads supported files and folders with adjustable concurrency and limits.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentBatch [-Path] <string[]> [-Recurse] [-Extension <string[]>] [-MaxDocuments <int>] [-NoDocumentLimit] [-MaxDegreeOfParallelism <int>] [-MaxStoreItems <int>] [-AllStoreItems] [-IncludePageLocations] [-ContinueOnError] [-Reader <OfficeDocumentReader>] [-ReaderOptions <ReaderOptions>] [-BatchOptions <ReaderBatchOptions>] [<CommonParameters>]
```

## DESCRIPTION
Reads supported files and folders with adjustable concurrency and limits.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeDocumentBatch -Path .\Reports -Recurse -MaxDegreeOfParallelism 4 -ContinueOnError
```

PSWriteOffice discovers registered formats and reports individual read failures without requiring .NET option objects.

## PARAMETERS

### -AllStoreItems
Project every matching item from each email store.

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

### -BatchOptions
Advanced batch settings supplied by a .NET host.

```yaml
Type: ReaderBatchOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ContinueOnError
Report individual read errors and continue processing other documents.

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
Optional extensions to include. Registered Reader formats are used automatically when omitted.

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

### -IncludePageLocations
Compute Word and RTF page locations when supported.

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

### -MaxDegreeOfParallelism
Maximum document reads in flight.

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

### -MaxDocuments
Maximum documents accepted in one batch. The default is 500.

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

### -MaxStoreItems
Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.

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

### -NoDocumentLimit
Remove the document-count ceiling.

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

### -Path
File, directory, or wildcard paths to read.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: FullName, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue, ByPropertyName)
Accept wildcard characters: True
```

### -Reader
Advanced immutable Reader configured by a .NET host.

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

### -ReaderOptions
Advanced source-reading settings supplied by a .NET host.

```yaml
Type: ReaderOptions
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Recurse
Search subdirectories when a path names a directory.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReadResult`

## RELATED LINKS

- None
