---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficePdfArchive
## SYNOPSIS
Converts a local DOC/DOCX/TXT directory to PDF with durable restart checks. Requires PowerShell 7.

## SYNTAX
### __AllParameterSets
```powershell
Export-OfficePdfArchive -InputDirectory <string> -OutputDirectory <string> -CheckpointDirectory <string> [-MaximumConcurrency <int>] [-MaximumFiles <int>] [-MaximumInputBytes <long>] [-MaximumOutputBytes <long>] [-TextEncoding <string>] [-TabSize <int>] [-MaximumTextCharacters <int>] [-MaximumTextPages <int>] [-RetryFailed] [-AllowLegacyImportLoss] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Source, output and checkpoint directories must be separate. Rerunning verifies completed source and output hashes.
Each output retains its full source filename plus .pdf. Known DOC import loss blocks output unless explicitly accepted.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficePdfArchive -InputDirectory .\Documents -OutputDirectory .\PDF -CheckpointDirectory .\PDF-State
```


## PARAMETERS

### -AllowLegacyImportLoss
Accept reported legacy DOC import loss.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CheckpointDirectory
Separate durable checkpoint tree.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputDirectory
Local DOC, DOCX and TXT source tree.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumConcurrency
Maximum simultaneously executing files.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumFiles
Maximum selected source files.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumInputBytes
Maximum source bytes per file.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumOutputBytes
Maximum PDF bytes per file.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTextCharacters
Maximum decoded and tab-expanded characters per TXT source.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumTextPages
Maximum generated pages per TXT source.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Separate PDF destination tree.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RetryFailed
Retry recorded failures, including corrected failed sources. Completed items remain protected.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TabSize
Literal text tab columns.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextEncoding
Explicit literal text encoding.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Workflows.OfficePdfArchiveResult`

## RELATED LINKS

- None
