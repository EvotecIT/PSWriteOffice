---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Search-OfficeDocument
## SYNOPSIS
Searches normalized document blocks and returns Reader-owned page citations for each match.

## SYNTAX
### __AllParameterSets
```powershell
Search-OfficeDocument [-InputObject] <OfficeDocumentReadResult> [-Query] <string> [-MatchCase] [-WholeWord] [-MaximumResults <int>] [<CommonParameters>]
```

## DESCRIPTION
Searches normalized document blocks and returns Reader-owned page citations for each match.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $document = Get-OfficeDocument -Path .\Policy.docx -IncludePageLocations
$matches = $document | Search-OfficeDocument -Query 'retention period'
$matches.Hits | Select-Object -ExpandProperty Pages
```

Uses OfficeIMO.Reader search and location contracts without reparsing document text in PowerShell.

## PARAMETERS

### -InputObject
Normalized document returned by Get-OfficeDocument.

```yaml
Type: OfficeDocumentReadResult
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -MaximumResults
Maximum number of occurrences to return.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Query
Text to find in normalized document blocks.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -WholeWord
Return only occurrences surrounded by non-word characters.

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

- `OfficeIMO.Reader.OfficeDocumentReadResult`

## OUTPUTS

- `OfficeIMO.Reader.OfficeDocumentSearchResult`

## RELATED LINKS

- None
