using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Copies selected PDF pages into a new PDF.</summary>
[Cmdlet(VerbsCommon.Copy, "OfficePdfPage")]
[OutputType(typeof(FileInfo))]
public sealed class CopyOfficePdfPageCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5.</summary>
    [Parameter(Mandatory = true)]
    public string PageRange { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).Pages.Extract(PageRange).Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
