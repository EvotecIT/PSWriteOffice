---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelTableOfContents
## SYNOPSIS
Adds or refreshes an Excel workbook table of contents sheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelTableOfContents [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelTableOfContents -Document <ExcelDocument> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelTableOfContents [-InputPath] <string> [-SheetName <string>] [-DoNotPlaceFirst] [-NoHyperlinks] [-IncludeNamedRanges] [-IncludeHiddenNamedRanges] [-NoStyle] [-AddBackLinks] [-BackLinkRow <int>] [-BackLinkColumn <int>] [-BackLinkText <string>] [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or refreshes a workbook navigation sheet based on the OfficeIMO Excel TOC helpers. It can be used inside `New-OfficeExcel`, with an open `ExcelDocument`, or directly against a workbook path.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeExcelTableOfContents -Path .\Report.xlsx -IncludeNamedRanges -AddBackLinks
```

Updates `Report.xlsx` in place with a TOC sheet and back links on other worksheets.

### EXAMPLE 2
```powershell
PS>New-OfficeExcel -Path .\Report.xlsx {
    ExcelSheet 'Data' { ExcelCell -Address 'A1' -Value 'Name' }
    ExcelTableOfContents
}
```

Creates a workbook and adds the TOC within the DSL using the `ExcelTableOfContents` alias.

## PARAMETERS

### -AddBackLinks
Add a link back to the TOC on each non-TOC sheet.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkColumn
Column used for backlinks when you want an explicit placement.

```yaml
Type: Int32
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkRow
Row used for backlinks when you want an explicit placement.

```yaml
Type: Int32
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: 2
Accept pipeline input: False
Accept wildcard characters: True
```

### -BackLinkText
Text shown for the backlink.

```yaml
Type: String
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: ← TOC
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Workbook to update.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -DoNotPlaceFirst
Keep the TOC sheet in its current position instead of moving it to the front.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -IncludeHiddenNamedRanges
Include hidden named ranges when `-IncludeNamedRanges` is used.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
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
Parameter Sets: Context, Document, Path
Aliases: None
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
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoStyle
Disable the styled TOC layout.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Open
Open the workbook after saving when `-Path` is used.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the updated workbook or file info.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document, Path
Aliases: None
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
Parameter Sets: Context, Document, Path
Aliases: None
Required: False
Position: named
Default value: TOC
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None
