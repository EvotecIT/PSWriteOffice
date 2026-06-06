using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a named bookmark at the current generated PDF flow position.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfBookmark", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfBookmark")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfBookmarkCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Bookmark name.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Name { get; set; } = string.Empty;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.Bookmark(Name);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
