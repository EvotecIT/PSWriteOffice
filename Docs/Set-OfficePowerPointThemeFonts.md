---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointThemeFonts
## SYNOPSIS
Sets PowerPoint theme fonts.

## SYNTAX
```powershell
Set-OfficePowerPointThemeFonts [-Presentation <PowerPointPresentation>] [-MajorLatin <string>] [-MinorLatin <string>] [-MajorEastAsian <string>] [-MinorEastAsian <string>] [-MajorComplexScript <string>] [-MinorComplexScript <string>] [-Master <int>] [-AllMasters] [-ClearMissing] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates one or more PowerPoint theme font slots. Unspecified font slots keep their existing values unless `-ClearMissing` is used.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointThemeFonts -Presentation $ppt -MajorLatin 'Aptos' -MinorLatin 'Calibri' -AllMasters
```

Applies Latin theme fonts across all masters.
