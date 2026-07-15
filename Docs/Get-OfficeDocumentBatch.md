---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeDocumentBatch
## SYNOPSIS
Reads a bounded set of documents asynchronously while retaining input order.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeDocumentBatch [-Path] <string[]> [-Reader <OfficeDocumentReader>] [-ReaderOptions <ReaderOptions>] [-BatchOptions <ReaderBatchOptions>] [<CommonParameters>]
```

## DESCRIPTION
Reads a bounded set of documents asynchronously while retaining input order.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $batch = [OfficeIMO.Reader.ReaderBatchOptions]::new(); $batch.MaxDegreeOfParallelism = 4; Get-ChildItem .\Reports -File | Get-OfficeDocumentBatch -BatchOptions $batch
```

OfficeIMO.Reader bounds concurrency and returns results in pipeline input order.

## PARAMETERS

### -BatchOptions
Optional maximum document count and degree of parallelism.

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

### -Path
Paths to read.

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
Optional immutable reader with caller-configured processors.

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
Optional source-reading limits and format behavior.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentReadResult`

## RELATED LINKS

- None
