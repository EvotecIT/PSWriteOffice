---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointTheme
## SYNOPSIS
Gets theme information for a PowerPoint presentation master.

## SYNTAX
```powershell
Get-OfficePowerPointTheme [-Presentation <PowerPointPresentation>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Returns a theme summary for a PowerPoint presentation master, including the theme name, defined theme colors, and configured theme fonts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointTheme -Presentation $ppt
```

Returns theme details for the default master.
