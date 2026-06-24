---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelThreadedComment
## SYNOPSIS
Adds a threaded comment or reply to an Excel worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelThreadedComment -Address <string> -Text <string> [-Author <string>] [-ParentId <string>] [-Id <string>] [-Date <datetime>] [-Done] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelThreadedComment [-InputPath] <string> -Address <string> -Text <string> [-Sheet <string>] [-SheetIndex <int>] [-Author <string>] [-ParentId <string>] [-Id <string>] [-Date <datetime>] [-Done] [-NoSave] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelThreadedComment -Document <ExcelDocument> -Address <string> -Text <string> [-Sheet <string>] [-SheetIndex <int>] [-Author <string>] [-ParentId <string>] [-Id <string>] [-Date <datetime>] [-Done] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Adds a threaded comment or reply to an Excel worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $comment = Add-OfficeExcelThreadedComment -Path .\Review.xlsx -Sheet Data -Address C5 -Text 'Please confirm this variance.' -Author 'Finance Reviewer' -PassThru
Add-OfficeExcelThreadedComment -Path .\Review.xlsx -Sheet Data -Address C5 -Text 'Confirmed with sales ops.' -Author 'Report Owner' -ParentId $comment.Id
Get-OfficeExcelCommentAudit -Path .\Review.xlsx -IncludeComments |
    Select-Object -ExpandProperty ThreadedComments
```

Uses OfficeIMO threaded-comment metadata authoring, including workbook person metadata, and keeps legacy notes separate.

## PARAMETERS

### -Address
A1-style cell address, such as C5.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Author
Author display name stored in workbook person metadata.

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

### -Date
Optional timestamp for the threaded comment.

```yaml
Type: Nullable`1
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
Workbook to operate on outside the DSL context.

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

### -Done
Mark the threaded comment as done/resolved.

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

### -Id
Optional stable threaded-comment id.

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

### -InputPath
Workbook path to update.

```yaml
Type: String
Parameter Sets: Path
Aliases: Path, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -NoSave
Do not save when operating on a path-owned workbook.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -ParentId
Optional parent threaded-comment id when adding a reply.

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

### -PassThru
Emit threaded-comment metadata.

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

### -Sheet
Worksheet name when using path or document input.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -SheetIndex
Worksheet index (0-based) when using path or document input.

```yaml
Type: Nullable`1
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### -Text
Threaded comment text.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: True
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None
