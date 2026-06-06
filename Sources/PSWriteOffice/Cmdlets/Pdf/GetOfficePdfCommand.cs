using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Opens an existing PDF as an OfficeIMO.Pdf document.</summary>
[Cmdlet(VerbsCommon.Get, "OfficePdf")]
[OutputType(typeof(PdfDocument))]
public sealed class GetOfficePdfCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        WriteObject(PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)));
    }
}
