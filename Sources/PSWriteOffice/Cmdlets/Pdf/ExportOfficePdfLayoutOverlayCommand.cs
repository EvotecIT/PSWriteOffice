using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports PDF word, line, region, and reading-order diagnostics as PNG or SVG.</summary>
[Cmdlet(VerbsData.Export, "OfficePdfLayoutOverlay", SupportsShouldProcess = true)]
[OutputType(typeof(OfficeImageExportResult))]
public sealed class ExportOfficePdfLayoutOverlayCommand : PSCmdlet
{
    /// <summary>Source PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination PNG or SVG path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>One-based page number.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int Page { get; set; } = 1;

    /// <summary>Output image format. Layout overlays support only PNG and SVG.</summary>
    [Parameter]
    [ValidateSet(nameof(OfficeImageExportFormat.Png), nameof(OfficeImageExportFormat.Svg))]
    public OfficeImageExportFormat Format { get; set; } = OfficeImageExportFormat.Svg;

    /// <summary>Output scale.</summary>
    [Parameter]
    [ValidateRange(0.01, 100.0)]
    public double Scale { get; set; } = 1.0;

    /// <summary>Optional overlay elements, colors, and limits.</summary>
    [Parameter]
    public PdfLayoutDebugOverlayOptions? Options { get; set; }

    /// <summary>Optional text layout settings.</summary>
    [Parameter]
    public PdfTextLayoutOptions? LayoutOptions { get; set; }

    /// <summary>Optional bounded PDF parsing settings.</summary>
    [Parameter]
    public PdfReadOptions? ReadOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var input = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        var output = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(output, $"Export PDF layout overlay as {Format}")) return;
        var drawing = PdfCommandUtilities.LoadDocument(input, ReadOptions).Read.LayoutDebugOverlay(Page, Options, LayoutOptions, ReadOptions);
        var bytes = Format switch
        {
            OfficeImageExportFormat.Svg => OfficeDrawingSvgExporter.ToSvgBytes(drawing, Scale, OfficeSvgSizeUnit.Pixel),
            OfficeImageExportFormat.Png => OfficeDrawingRasterRenderer.ToPng(drawing, Scale),
            _ => throw new PSArgumentException("PDF layout overlays support only PNG and SVG output.", nameof(Format))
        };
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(output) ?? SessionState.Path.CurrentFileSystemLocation.Path);
        File.WriteAllBytes(output, bytes);
        WriteObject(new OfficeImageExportResult(Format,
            checked((int)System.Math.Ceiling(drawing.Width * Scale)),
            checked((int)System.Math.Ceiling(drawing.Height * Scale)),
            bytes, $"Page {Page} layout", $"{input}#page={Page}"));
    }
}
