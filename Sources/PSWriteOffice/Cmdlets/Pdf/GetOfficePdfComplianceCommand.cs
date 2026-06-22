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
[Cmdlet(VerbsCommon.Get, "OfficePdfCompliance", DefaultParameterSetName = ParameterSetDocument)]
[OutputType(typeof(PdfComplianceReadinessReport))]
public sealed class GetOfficePdfComplianceCommand : PSCmdlet
{
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Generated PDF document to assess outside the DSL context.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Existing PDF file path to assess after generation.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Compliance profile to assess. When omitted, the document's configured profile is used.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPath)]
    public PdfComplianceProfile? Profile { get; set; }

    /// <summary>Password used to inspect a Standard password-encrypted PDF.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetPath)
        {
            WriteObject(PdfComplianceAnalyzer.AssessReadback(Profile!.Value, PdfCommandUtilities.ResolvePath(this, Path), PdfCommandUtilities.CreateReadOptions(Password)));
            return;
        }

        var document = Document ?? PdfDslContext.Require(this).Document;
        WriteObject(Profile.HasValue ? document.AssessCompliance(Profile.Value) : document.AssessCompliance());
    }
}
