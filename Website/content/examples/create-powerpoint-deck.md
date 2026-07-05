---
title: "Create a PowerPoint deck"
description: "Use PSWriteOffice to generate a PowerPoint deck with slides, sections, notes, and transitions."
layout: docs
---

This pattern is useful when a repeatable status script should produce a deck that can still be edited in PowerPoint.

It is adapted from `Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1` and `Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1`.

## Example

```powershell
Import-Module PSWriteOffice

$outputPath = Join-Path $PSScriptRoot 'Output\ServiceBrief.pptx'
New-Item -ItemType Directory -Path (Split-Path $outputPath) -Force | Out-Null

$deck = New-OfficePowerPoint -Path $outputPath {
    PptSlide {
        PptTitle -Title 'Service Brief'
        PptTextBox -Text 'Generated with PSWriteOffice' -X 80 -Y 145 -Width 360 -Height 50
        PptBullets -Bullets 'Current state', 'Risks', 'Next steps' -X 80 -Y 220 -Width 520 -Height 180
        PptNotes -Text 'Use this slide as the spoken summary.'
    }
}

Add-OfficePowerPointSection -Presentation $deck -Name 'Opening' -StartSlideIndex 0
Get-OfficePowerPointSlide -Presentation $deck -Index 0 |
    Set-OfficePowerPointSlideTransition -Transition Fade

Set-OfficePowerPointSlideSize -Presentation $deck -Preset Screen16x9
Save-OfficePowerPoint -Presentation $deck
```

## What this demonstrates

- creating an editable `.pptx` from PowerShell
- adding slide content, speaker notes, and sections
- applying deck-level sizing and slide transitions

## Source

- [Example-PowerPointTransitionsAndSizing.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Example-PowerPointTransitionsAndSizing.ps1)
- [Example-PowerPointSectionsAndImport.ps1](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Example-PowerPointSectionsAndImport.ps1)
