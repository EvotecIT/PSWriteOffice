using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports whether OfficeIMO.Pdf can read or rewrite a PDF safely.</summary>
/// <example>
///   <summary>Preflight a PDF before migration operations.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$preflight = Get-OfficePdfPreflight -Path .\Examples\Documents\Report.pdf
/// $preflight.HasReadBlockers
/// $preflight.HasRewriteBlockers</code>
///   <para>Checks whether OfficeIMO.Pdf can read or rewrite the PDF safely.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfPreflight")]
[OutputType(typeof(PdfDocumentPreflight))]
public sealed class GetOfficePdfPreflightCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfInspector.Preflight(PdfCommandUtilities.ResolvePath(this, Path)));
    }
}
