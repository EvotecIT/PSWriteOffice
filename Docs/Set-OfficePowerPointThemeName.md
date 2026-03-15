---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointThemeName
## SYNOPSIS
Sets the PowerPoint theme name.

## SYNTAX
```powershell
Set-OfficePowerPointThemeName [-Presentation <PowerPointPresentation>] -Name <string> [-AllMasters] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates the PowerPoint theme name. Use `-AllMasters` to apply the name across every master in the presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointThemeName -Presentation $ppt -Name 'Contoso Theme' -AllMasters
```

Renames the theme across the presentation.
