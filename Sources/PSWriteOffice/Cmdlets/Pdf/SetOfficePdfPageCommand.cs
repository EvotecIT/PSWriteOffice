using System;
using System.IO;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets page-level PDF properties and writes a new PDF.</summary>
/// <example>
///   <summary>Rotate selected PDF pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Set-OfficePdfPage -Path .\Examples\Documents\Scanned.pdf -PageRange '2,4' -Rotation 90 -OutputPath .\Examples\Documents\Scanned-Rotated.pdf
///     Get-OfficePdfInfo -Path .\Examples\Documents\Scanned-Rotated.pdf | Select-Object PageCount
/// )
/// $proof</code>
///   <para>Rotates selected pages and writes a new PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfPage", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class SetOfficePdfPageCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5. Omit to affect all pages.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Rotation in degrees. Supported values are 0, 90, 180, and 270.</summary>
    [Parameter]
    [ValidateSet("0", "90", "180", "270")]
    public int Rotation { get; set; }

    /// <summary>Page boundary box to set. Supported values are MediaBox, CropBox, BleedBox, TrimBox, and ArtBox.</summary>
    [Parameter]
    [ValidateSet("MediaBox", "CropBox", "BleedBox", "TrimBox", "ArtBox")]
    public string? BoxName { get; set; }

    /// <summary>Left coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Left { get; set; }

    /// <summary>Bottom coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Bottom { get; set; }

    /// <summary>Right coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Right { get; set; }

    /// <summary>Top coordinate for the page boundary box.</summary>
    [Parameter]
    public double? Top { get; set; }

    /// <summary>Resize selected pages to a known OfficeIMO page size such as A4, Letter, or Custom.</summary>
    [Parameter]
    public string? PageSize { get; set; }

    /// <summary>Custom page width in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Custom page height in points when -PageSize Custom is used.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Use the landscape orientation of the selected page size.</summary>
    [Parameter]
    public SwitchParameter Landscape { get; set; }

    /// <summary>How source page content is fitted into the resized output page.</summary>
    [Parameter]
    public PdfPageResizeMode ResizeMode { get; set; } = PdfPageResizeMode.Fit;

    /// <summary>Margin, in points, reserved around resized page content.</summary>
    [Parameter]
    public double? ResizeMargin { get; set; }

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write updated PDF pages"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        int[] pages = string.IsNullOrWhiteSpace(PageRange)
            ? Array.Empty<int>()
            : PdfPageRange.ParseMany(PageRange!).SelectMany(ExpandPageRange).Distinct().ToArray();
        var resizeOptions = PdfCommandUtilities.CreatePageResizeOptions(
            PageSize,
            Width,
            Height,
            Landscape.IsPresent,
            ResizeMode,
            ResizeMargin,
            MyInvocation.BoundParameters.ContainsKey(nameof(PageSize)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Width)) ||
            MyInvocation.BoundParameters.ContainsKey(nameof(Height)) ||
            Landscape.IsPresent ||
            MyInvocation.BoundParameters.ContainsKey(nameof(ResizeMode)) ||
            ResizeMargin.HasValue);

        if (resizeOptions != null)
        {
            if (!string.IsNullOrWhiteSpace(BoxName) || MyInvocation.BoundParameters.ContainsKey(nameof(Rotation)))
            {
                throw new PSArgumentException("Use page resize, rotation, or box editing as separate Set-OfficePdfPage operations.");
            }

            PdfPageEditor.ResizePages(inputPath, outputPath, resizeOptions, pages);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        if (!string.IsNullOrWhiteSpace(BoxName))
        {
            if (!Left.HasValue || !Bottom.HasValue || !Right.HasValue || !Top.HasValue)
            {
                throw new PSArgumentException("-BoxName requires -Left, -Bottom, -Right, and -Top.");
            }

            PdfPageEditor.SetPageBox(inputPath, outputPath, BoxName!, Left.Value, Bottom.Value, Right.Value, Top.Value, pages);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        if (!MyInvocation.BoundParameters.ContainsKey(nameof(Rotation)))
        {
            throw new PSArgumentException("Provide -Rotation, -BoxName with coordinates, or page resize options.");
        }

        var document = PdfDocument.Load(inputPath);
        var result = string.IsNullOrWhiteSpace(PageRange)
            ? document.Pages.Rotate(Rotation)
            : document.Pages.Rotate(Rotation, PageRange!);
        result.Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }

    private static int[] ExpandPageRange(PdfPageRange range)
    {
        return Enumerable.Range(range.FirstPage, range.PageCount).ToArray();
    }
}
