using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Clears generated PDF page background shapes.</summary>
/// <example>
///   <summary>Remove background shapes before saving a variant.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = New-OfficePdf {
///     Add-OfficePdfBackgroundShape -Shape Rectangle -Color '#EEF2FF' -X 0 -Y 0 -Width 595 -Height 120
///     Add-OfficePdfHeading -Text 'Clean variant'
/// } -NoSave
/// $pdf | Clear-OfficePdfBackgroundShape | Save-OfficePdf -Path .\Examples\Documents\PdfNoBackgroundShape.pdf</code>
///   <para>Clears generated page background shapes on an in-memory PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Clear, "OfficePdfBackgroundShape", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(PdfDocument))]
public sealed class ClearOfficePdfBackgroundShapeCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.ClearBackgroundShapes();
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
