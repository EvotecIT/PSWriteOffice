using System.IO;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets or clears a generated PDF page background image.</summary>
/// <remarks>
/// Background images are applied through OfficeIMO.Pdf and rendered behind normal generated content.
/// Use low opacity for watermark-like page texture and <c>-Clear</c> to remove a previously configured background image.
/// </remarks>
/// <example>
///   <summary>Add a subtle background image.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfBackgroundImage -Path .\letterhead.png -Fit Cover -Opacity 0.08
///     PdfHeading 'Branded report'
///   }</code>
///   <para>Uses an image as a low-opacity page background.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfBackgroundImage", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBackgroundImage")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfBackgroundImageCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Background image path.</summary>
    [Parameter(Position = 0)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>How the image should fit the page box.</summary>
    [Parameter]
    public OfficeImageFit Fit { get; set; } = OfficeImageFit.Cover;

    /// <summary>Image opacity from 0 to 1.</summary>
    [Parameter]
    [ValidateRange(0D, 1D)]
    public double? Opacity { get; set; }

    /// <summary>Clear the generated PDF background image.</summary>
    [Parameter]
    public SwitchParameter Clear { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        if (Clear.IsPresent)
        {
            document.BackgroundImage(null);
        }
        else
        {
            if (string.IsNullOrWhiteSpace(Path))
            {
                throw new PSArgumentException("-Path is required unless -Clear is used.", nameof(Path));
            }

            var imagePath = PdfCommandUtilities.ResolvePath(this, Path!);
            document.BackgroundImage(File.ReadAllBytes(imagePath), Fit, Opacity);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
