---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordPageSetup
## SYNOPSIS
Sets page setup options on Word sections.

## SYNTAX
### Current (Default)
```powershell
Set-OfficeWordPageSetup [-PageSize <WordPageSize>] [-Orientation <string>] [-Margin <WordMargin>] [-Left <int>] [-Right <int>] [-Top <int>] [-Bottom <int>] [-Header <int>] [-Footer <int>] [-Gutter <int>] [-Columns <int>] [-ColumnSpacing <int>] [-ColumnSeparator <bool>] [-PassThru] [<CommonParameters>]
```

### Section
```powershell
Set-OfficeWordPageSetup -Section <WordSection> [-PageSize <WordPageSize>] [-Orientation <string>] [-Margin <WordMargin>] [-Left <int>] [-Right <int>] [-Top <int>] [-Bottom <int>] [-Header <int>] [-Footer <int>] [-Gutter <int>] [-Columns <int>] [-ColumnSpacing <int>] [-ColumnSeparator <bool>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeWordPageSetup -Document <WordDocument> [-Index <int[]>] [-PageSize <WordPageSize>] [-Orientation <string>] [-Margin <WordMargin>] [-Left <int>] [-Right <int>] [-Top <int>] [-Bottom <int>] [-Header <int>] [-Footer <int>] [-Gutter <int>] [-Columns <int>] [-ColumnSpacing <int>] [-ColumnSeparator <bool>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates page size, orientation, margins, and section columns through OfficeIMO.Word.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordSection { Set-OfficeWordPageSetup -PageSize A4 -Orientation Landscape -Columns 2 }
```

Updates the current section page setup.

## PARAMETERS

### -Bottom
Bottom margin in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Columns
Number of section columns.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnSeparator
Whether to show a separator between columns.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ColumnSpacing
Space between columns in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Document
Document whose sections should be updated.

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

### -Footer
Footer distance in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Gutter
Gutter size in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Header
Header distance in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Index
Optional 0-based section indexes when -Document is used.

```yaml
Type: Int32[]
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Left
Left margin in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Margin
Built-in margin preset.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Orientation
Page orientation.

```yaml
Type: String
Parameter Sets: Current, Section, Document
Aliases: None
Possible values: Portrait, Landscape

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PageSize
Built-in page size.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit updated sections.

```yaml
Type: SwitchParameter
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Right
Right margin in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Section
Section to update.

```yaml
Type: WordSection
Parameter Sets: Section
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Top
Top margin in twips.

```yaml
Type: Nullable`1
Parameter Sets: Current, Section, Document
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

- `OfficeIMO.Word.WordSection
OfficeIMO.Word.WordDocument`

## OUTPUTS

- `OfficeIMO.Word.WordSection`

## RELATED LINKS

- None
