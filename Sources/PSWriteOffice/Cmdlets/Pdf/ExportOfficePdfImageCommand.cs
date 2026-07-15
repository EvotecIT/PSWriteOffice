using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports PDF pages as PNG or SVG and normalizes each page to OfficeImageExportResult.</summary>
/// <example>
///   <summary>Export selected pages as PNG files.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\Pages -PageRange '1-3,5'</code>
///   <para>Writes the selected pages and returns normalized image results with rendering diagnostics.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficePdfImage", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficePdfImageCommand : PSCmdlet
{
    /// <summary>Path to the PDF.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination folder.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Optional one-based ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Output image format.</summary>
    [Parameter]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Png;

    /// <summary>Optional DPI, scale, thumbnail, limits, and error behavior.</summary>
    [Parameter]
    public PdfPageRenderOptions? Options { get; set; }

    /// <summary>Optional bounded PDF parsing settings.</summary>
    [Parameter]
    public PdfReadOptions? ReadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export PDF pages as {Format}")) return;
        Directory.CreateDirectory(output);
        var sourceOptions = Options ?? new PdfPageRenderOptions();
        var options = new PdfPageRenderOptions
        {
            Format = Format == OfficeImageExportFormat.Svg ? PdfPageRenderFormat.Svg : PdfPageRenderFormat.Png,
            Scale = sourceOptions.Scale,
            Dpi = sourceOptions.Dpi,
            Background = sourceOptions.Background,
            ThumbnailMaxDimension = sourceOptions.ThumbnailMaxDimension,
            MaxPages = sourceOptions.MaxPages,
            MaxPixelsPerPage = sourceOptions.MaxPixelsPerPage,
            ContinueOnError = sourceOptions.ContinueOnError,
            ImageCodec = sourceOptions.ImageCodec
        };
        var document = PdfCommandUtilities.LoadDocument(input, ReadOptions);
        IReadOnlyList<PdfPageRenderResult> pages = string.IsNullOrWhiteSpace(PageRange)
            ? document.Read.RenderPages(options: options, readOptions: ReadOptions)
            : document.Read.RenderPages(PageRange!, options, ReadOptions);
        var extension = Format == OfficeImageExportFormat.Svg ? ".svg" : ".png";
        foreach (var page in pages)
        {
            if (!page.Succeeded || page.Bytes == null)
            {
                WriteError(new ErrorRecord(
                    new InvalidDataException($"PDF page {page.PageNumber} could not be rendered: {string.Join("; ", page.Diagnostics)}"),
                    "OfficePdfImageExportFailed", ErrorCategory.InvalidData, page.PageNumber));
                continue;
            }
            var file = System.IO.Path.Combine(output, $"page-{page.PageNumber:D4}{extension}");
            File.WriteAllBytes(file, page.Bytes);
            var diagnostics = page.Diagnostics.Select(message => new OfficeImageExportDiagnostic(
                OfficeImageExportDiagnosticSeverity.Warning, "PdfRenderDiagnostic", message, $"page {page.PageNumber}"))
                .ToArray();
            WriteObject(new OfficeImageExportResult(
                Format, page.Width, page.Height, page.Bytes,
                $"Page {page.PageNumber}", $"{input}#page={page.PageNumber}", diagnostics));
        }
    }
}
