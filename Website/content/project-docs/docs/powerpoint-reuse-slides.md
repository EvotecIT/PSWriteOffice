---
title: "Reuse and Combine PowerPoint Slides"
description: "Copy slides inside a deck, import slides from another presentation, and organize combined material into sections."
layout: docs
---

Slide reuse is the practical PowerPoint equivalent of document merge. It supports reusable architecture slides, standard legal pages, customer-specific inserts, and decks assembled from team-owned sources.

## Copy or import

- `Copy-OfficePowerPointSlide` duplicates a slide inside the current presentation.
- `Import-OfficePowerPointSlide` imports a selected slide from another presentation path or loaded presentation.
- `Add-OfficePowerPointSection`, `Rename-OfficePowerPointSection`, and related commands organize the combined deck.

```powershell
Import-OfficePowerPointSlide -Presentation $target `
    -SourcePath '.\Reusable-Slides.pptx' -SourceIndex 0 -InsertAt 2
```

Inspect imported theme and layout behavior when the source and target use different design systems. Reopen the saved deck and review slide summaries, placeholders, and notes.

The [reuse-slides recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/PowerPoint/Recipe-PowerPoint-ReuseSlides.ps1) imports one slide, copies another, creates a section, and reports the final counts.
