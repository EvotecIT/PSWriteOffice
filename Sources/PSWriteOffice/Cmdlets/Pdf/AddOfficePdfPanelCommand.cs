using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Adds a visually separated panel paragraph to a PDF document.</summary>
[Cmdlet(VerbsCommon.Add, "OfficePdfPanel", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfPanel")]
[OutputType(typeof(PdfDocument))]
public sealed class AddOfficePdfPanelCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Panel text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Panel alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Left;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.PanelParagraph(p => p.Text(Text), align: Align);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
