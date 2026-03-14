---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordField
## SYNOPSIS
Adds a field to the current paragraph.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeWordField [-Type] <WordFieldType> [-Format <WordFieldFormat>] [-CustomFormat <string>] [-Advanced] [-Parameters <string[]>] [-Paragraph <WordParagraph>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a field to the current paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Add-OfficeWordParagraph { Add-OfficeWordField -Type Page }
```

Inserts a PAGE field into the paragraph.

## PARAMETERS

### -Advanced
Use advanced field representation.

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

### -CustomFormat
Custom format string (date/time fields).

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

### -Format
Optional field format switch.

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

### -Paragraph
Explicit paragraph to receive the field.

```yaml
Type: WordParagraph
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Parameters
Additional field parameters.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -PassThru
Emit the paragraph after adding the field.

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

### -Type
Field type to insert.

```yaml
Type: WordFieldType
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: AddressBlock, Advance, Ask, Author, AutoNum, AutoNumLgl, AutoNumOut, AutoText, AutoTextList, Bibliography, Citation, Comments, Compare, CreateDate, Database, Date, DocProperty, DocVariable, Embed, FileName, FileSize, GoToButton, GreetingLine, HyperlinkIf, IncludePicture, IncludeText, Index, Info, Keywords, LastSavedBy, Link, ListNum, MacroButton, MergeField, MergeRec, MergeSeq, Next, NextIf, NoteRef, NumChars, NumPages, NumWords, Page, PageRef, Print, PrintDate, Private, Quote, RD, Ref, RevNum, SaveDate, Section, SectionPages, Seq, Set, SkipIf, StyleRef, Subject, Symbol, TA, TC, Template, Time, Title, TOA, TOC, UserAddress, UserInitials, UserName, XE

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `System.Object`

## RELATED LINKS

- None

