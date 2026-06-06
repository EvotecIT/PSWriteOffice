using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds an image to a PDF document.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfImage", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfImage")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfImageCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Image path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Rendered image width in PDF points.</summary>
    [Parameter(Mandatory = true)]
    public double Width { get; set; }

    /// <summary>Rendered image height in PDF points.</summary>
    [Parameter(Mandatory = true)]
    public double Height { get; set; }

    /// <summary>Image alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Alternative text for meaningful images.</summary>
    [Parameter]
    public string? AlternativeText { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        var imagePath = PdfCommandUtilities.ResolvePath(this, Path);
        document.Image(File.ReadAllBytes(imagePath), Width, Height, Align, null, null, null, null, null, null, null, AlternativeText);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
