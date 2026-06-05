using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Moves selected pages before another page and writes a new PDF.</summary>
[Cmdlet(VerbsCommon.Move, "OfficePdfPage")]
[OutputType(typeof(FileInfo))]
public sealed class MoveOfficePdfPageCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5.</summary>
    [Parameter(Mandatory = true)]
    public string PageRange { get; set; } = string.Empty;

    /// <summary>One-based page number before which selected pages are inserted. Use page count + 1 to move to the end.</summary>
    [Parameter(Mandatory = true)]
    public int BeforePage { get; set; }

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path)).Pages.Move(BeforePage, PageRange).Save(outputPath);
        WriteObject(new FileInfo(outputPath));
    }
}
