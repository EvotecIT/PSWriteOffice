---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideLayout
## SYNOPSIS
Changes the layout used by a slide.

## SYNTAX
### Index (Default)
```powershell
Set-OfficePowerPointSlideLayout [-Slide <PowerPointSlide>] [-Master <int>] -Layout <int> [<CommonParameters>]
```

### Name
```powershell
Set-OfficePowerPointSlideLayout [-Slide <PowerPointSlide>] [-Master <int>] -LayoutName <string> [-CaseSensitive] [<CommonParameters>]
```

### Type
```powershell
Set-OfficePowerPointSlideLayout [-Slide <PowerPointSlide>] [-Master <int>] -LayoutType <SlideLayoutValues> [<CommonParameters>]
```

## DESCRIPTION
Changes the layout used by an existing slide using a layout index, layout name, or layout type from the selected master.

## EXAMPLES

### EXAMPLE 1
```powershell
PS>Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideLayout -LayoutName 'Title and Content'
```

Switches the first slide to the requested layout.
