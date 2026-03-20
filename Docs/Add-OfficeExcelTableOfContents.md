---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTableOfContents
## SYNOPSIS
Adds or refreshes a workbook table of contents sheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelTableOfContents [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelTableOfContents [-InputPath] <string> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelTableOfContents -Document <ExcelDocument> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or refreshes a workbook table of contents sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeExcelTableOfContents -Path .\report.xlsx -IncludeNamedRanges -AddBackLinks
```

Creates or refreshes a TOC sheet, lists named ranges, and adds back links on other sheets.

## PARAMETERS

### -AddBackLinks
Add a quick link back to the TOC on each worksheet.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkColumn
Column for the back link when AddBackLinks is used.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkRow
Row for the back link when AddBackLinks is used.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkText
Text used for back links.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to update.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values: 

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -DoNotPlaceFirst
Keep the TOC sheet in its current position instead of moving it first.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeHiddenNamedRanges
Include hidden named ranges when listing named ranges.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeNamedRanges
Include named ranges in the TOC.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -InputPath
Path to the workbook to update in place.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath, Path
Possible values: 

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoHyperlinks
Disable internal hyperlinks in the TOC sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoStyle
Disable formatted TOC styling.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Open
Open the workbook after saving when InputPath is used.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated document or file info.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetName
Name of the TOC sheet.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument
System.IO.FileInfo`

## RELATED LINKS

- None

