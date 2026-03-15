---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointThemeColor
## SYNOPSIS
Sets one or more PowerPoint theme colors.

## SYNTAX
### Single (Default)
```powershell
Set-OfficePowerPointThemeColor [-Presentation <PowerPointPresentation>] -Color <PowerPointThemeColor> -Value <string> [-Master <int>] [-AllMasters] [-PassThru] [<CommonParameters>]
```

### Multiple
```powershell
Set-OfficePowerPointThemeColor [-Presentation <PowerPointPresentation>] -Colors <hashtable> [-Master <int>] [-AllMasters] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates PowerPoint theme colors using single-color or hashtable-based input. Use `-AllMasters` to apply the same palette across every slide master.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Set-OfficePowerPointThemeColor -Presentation $ppt -Color Accent1 -Value '#C00000'
```

Updates Accent1 on the default master.

### EXAMPLE 2
```powershell
PS>Set-OfficePowerPointThemeColor -Presentation $ppt -Colors @{ Accent1 = '#C00000'; Accent2 = '00B0F0' } -AllMasters
```

Applies multiple theme color changes across all masters.
