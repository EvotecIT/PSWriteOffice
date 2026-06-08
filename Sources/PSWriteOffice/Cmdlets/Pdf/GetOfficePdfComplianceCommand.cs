using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets a generated PDF document compliance readiness report.</summary>
/// <example>
///   <summary>Check generated PDF compliance readiness.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$pdf = New-OfficePdf {
///     Set-OfficePdfCompliance -Profile PdfA3B -Groundwork
///     Add-OfficePdfHeading -Text 'Compliance readiness'
/// } -NoSave
/// $pdf | Get-OfficePdfCompliance -Profile PdfA3B</code>
///   <para>Returns the OfficeIMO.Pdf readiness report before saving.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfCompliance", DefaultParameterSetName = ParameterSetContext)]
[OutputType(typeof(PdfComplianceReadinessReport))]
public sealed class GetOfficePdfComplianceCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Generated PDF document to assess outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Compliance profile to assess. When omitted, the document's configured profile is used.</summary>
    [Parameter]
    public PdfComplianceProfile? Profile { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        WriteObject(Profile.HasValue ? document.AssessCompliance(Profile.Value) : document.AssessCompliance());
    }
}
