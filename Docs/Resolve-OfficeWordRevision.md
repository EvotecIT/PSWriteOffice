---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Resolve-OfficeWordRevision
## SYNOPSIS
Accepts or rejects filtered Word revisions and returns an operation report.

## SYNTAX
### Path (Default)
```powershell
Resolve-OfficeWordRevision [-Path] <string> -Action <WordRevisionOperationKind> -OutputPath <string> [-Filter <WordRevisionFilter>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Resolve-OfficeWordRevision -Document <WordDocument> -Action <WordRevisionOperationKind> [-Filter <WordRevisionFilter>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Accepts or rejects filtered Word revisions and returns an operation report.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $filter = [OfficeIMO.Word.WordRevisionFilter]::new(); $filter.Author = 'Reviewer'; Resolve-OfficeWordRevision -Path .\Draft.docx -OutputPath .\Accepted.docx -Action Accept -Filter $filter
```

Applies only matching revisions, saves the result, and returns the matched revision report.

## PARAMETERS

### -Action
Accept or reject matching revisions.

```yaml
Type: WordRevisionOperationKind
Parameter Sets: Path, Document
Aliases: None
Possible values: Accept, Reject

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Open Word document instance.

```yaml
Type: WordDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Filter
Optional author, id, type, date, location, part, or container filter.

```yaml
Type: WordRevisionFilter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -OutputPath
Output document path. Required for path input.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Return the mutated document after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Path
Path to the Word document.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordRevisionOperationReport`

## RELATED LINKS

- None
