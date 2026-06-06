using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes selected pages from a PDF and writes a new PDF.</summary>
/// <example>
///   <summary>Remove draft pages from a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficePdfPage -Path .\Examples\Documents\Report.pdf -PageRange '4-5' -OutputPath .\Examples\Documents\Report-Clean.pdf
/// Get-OfficePdfPreflight -Path .\Examples\Documents\Report-Clean.pdf</code>
///   <para>Deletes selected pages and preflights the rewritten PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficePdfPage")]
[OutputType(typeof(FileInfo))]
public sealed class RemoveOfficePdfPageCommand : PSCmdlet
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
        PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).Pages.Delete(PageRange).Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
