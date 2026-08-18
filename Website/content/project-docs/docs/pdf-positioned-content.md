---
title: "Position Text and Graphics on PDF Pages"
description: "Choose between flowing PDF text, positioned stamps, arbitrary page canvases, and imported page overlays with explicit coordinates."
layout: docs
---

A `PdfText` run belongs to normal document flow. It can change emphasis, color, font, baseline, and links, but it does not independently choose a starting coordinate. Fixed positioning is a page-level operation.

## Choose the positioning surface

- Use `PdfParagraph`, `PdfText`, and rich runs when content should reflow naturally.
- Use `Add-OfficePdfStamp -X -Y` for one positioned text or image stamp on an existing PDF.
- Use `Add-OfficePdfCanvas` for arbitrary text, rich text, images, shapes, drawings, or tables at page-aware coordinates.
- Use `Add-OfficePdfPageOverlay` to place an imported PDF page in a target rectangle, optionally as an underlay.

## Draw rich text at a coordinate

Canvas coordinates use PDF points from the visual top-left of the page. The callback runs once per selected page and exposes its width, height, page number, and page count.

```powershell
$runs = @(
    New-OfficeTextRun -Text 'Owner: ' -Bold | ConvertTo-OfficePdfTextRun
    New-OfficeTextRun -Text 'Platform' -Color '#0F766E' | ConvertTo-OfficePdfTextRun
)
$nativeRuns = [Array]::CreateInstance($runs[0].GetType(), $runs.Count)
for ($index = 0; $index -lt $runs.Count; $index++) {
    $nativeRuns.SetValue($runs[$index], $index)
}

Add-OfficePdfCanvas -Path '.\Report.pdf' -OutputPath '.\Positioned.pdf' -Content {
    param($canvas, $page)
    $null = $canvas.Text($nativeRuns, 36, 24, $page.Width - 72, 24, $null, 'Left', 10, 12)
}
```

The [positioned-canvas recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-PositionedCanvas.ps1) adds a rich review header and page-aware footer to a two-page document.

This coordinate system differs from lower-level PDF editing and redaction APIs that expose native bottom-left PDF coordinates. Follow the coordinate contract documented by the command being used rather than mixing values between surfaces.
