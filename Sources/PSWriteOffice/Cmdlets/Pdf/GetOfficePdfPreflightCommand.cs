using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reports whether OfficeIMO.Pdf can read or rewrite a PDF safely.</summary>
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
