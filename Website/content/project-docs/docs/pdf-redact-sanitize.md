---
title: "Redact, Sanitize, and Optimize PDF Files"
description: "Remove sensitive content by detected coordinates, sanitize interactive structures, and optimize a verified delivery copy."
layout: docs
---

Redaction, sanitization, and optimization are separate operations. Redaction removes selected content. Sanitization removes or normalizes risky document structures according to policy. Optimization changes representation to reduce or simplify the file.

## Redact detected text

Read text blocks with `Get-OfficePdfText -AsTextBlock`, select the intended block, calculate a bounded rectangle, and apply `ConvertTo-OfficePdfRedacted`. Re-extract text from the result to verify that the secret is gone and surrounding content remains.

The [redact-detected-text recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-RedactDetectedText.ps1) demonstrates that complete loop.

## Prepare a delivery copy

`ConvertTo-OfficePdfSanitized` writes a new sanitized file. `ConvertTo-OfficePdfOptimized -PassThruReport` reports the actions and size outcome. Neither should overwrite the only source copy in an auditable workflow.

The [sanitize-and-optimize recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Pdf/Recipe-Pdf-SanitizeAndOptimize.ps1) preserves each stage, captures size evidence, and preflights the final output.

Visual black boxes are not automatically secure redactions. Use the redaction command and verify extracted content rather than relying on a stamp or shape to hide text.
