---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Split-OfficePdf
## SYNOPSIS
Splits a PDF into page, range, count, or bookmark files.

## SYNTAX
### __AllParameterSets
```powershell
Split-OfficePdf [-Path] <string> [-OutputDirectory] <string> [-Prefix <string>] [-PagesPerDocument <int>] [-PageRange <string[]>] [-BookmarkName <string[]>] [-ByBookmark] [-Password <string>] [-PadIndex] [-IndexWidth <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Splits a PDF into page, range, count, or bookmark files.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $pages = Split-OfficePdf -Path .\Examples\Documents\Combined.pdf -OutputDirectory .\Examples\Documents\Pages -Prefix 'combined-page'
$pages | Select-Object Name, Length
```

Creates one output PDF for each page and returns the written files.

## PARAMETERS

### -BookmarkName
Create one PDF for each supplied bookmark title.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: Bookmark, BookmarkTitle
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ByBookmark
Create one PDF for every readable bookmark when -BookmarkName is not supplied.

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

### -IndexWidth
Pad numeric split names to this explicit width.

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
Output directory.

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

### -PadIndex
Pad numeric split names to the width required by the source page count.

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

### -PageRange
Create one PDF for each supplied page range or selection, such as 1-3 or 1,3.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: Range, PageRanges
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PagesPerDocument
Create one PDF for each consecutive group with this many pages.

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

### -Password
Password used to open an encrypted PDF.

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
Input PDF path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Prefix
Output file prefix.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None
