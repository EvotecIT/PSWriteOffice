# PSWriteOffice 3.0.8 (unreleased)

This release updates all OfficeIMO NuGet dependencies from 3.4.3 to 3.4.4. PowerShell 5.1 and PowerShell 7 remain supported.

Existing commands gain these improvements from the document engines:

- Excel saves preserve carriage returns and CRLF in cell text. XLSX and CSV writing also receive allocation and throughput improvements.
- PDF conversion improves text wrapping, embedded-font ligatures, glyph rendering, and positioned Word tables. PDF reading, extraction, mutation, and rendering use bounded processing for untrusted inputs.
- OpenDocument conversion preserves more spreadsheet validations, conditional styles, charts, and supported pivots; Word footnotes, fields, and page-specific headers; and PowerPoint presentation structure. Unsupported content remains subject to conversion-loss reporting.
- Word, PowerPoint, and Visio document workflows receive memory and performance improvements through their normal load/save APIs.

PDF image composition, table-cell images, backgrounds, stamps, and watermarks now reject encoded image files larger than OfficeIMO's default 128 MiB limit before buffering the entire file. Electronic-invoice configuration uses OfficeIMO's bounded XML file reader and rejects invoice XML larger than 16 MiB. Oversized inputs fail without producing the requested PDF.

`Get-OfficeProtectionCapability -AsJson` now returns schema version 2 from OfficeIMO. Consumers should read the operation dispositions and `unsupportedOperations` fields rather than assuming the previous catalog shape.

The upstream changes are listed in the [OfficeIMO September 27 release](https://github.com/EvotecIT/OfficeIMO/releases/tag/OfficeIMO-v20260927150127).

## Further PowerShell coverage

The dependency update does not add commands for every new engine API. Useful candidates for separate work are reviewed static-PDF form recognition, booklet/N-up imposition, richer PDF annotation authoring, and Project document automation. These need PowerShell input, output, overwrite, and review contracts before becoming public commands. Existing OpenDocument conversion commands already consume the improved adapters; direct ODS chart and pivot authoring could be exposed separately if native ODS editing is needed.
