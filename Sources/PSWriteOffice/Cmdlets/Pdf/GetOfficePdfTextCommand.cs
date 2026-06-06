using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Extracts text or Markdown from a PDF.</summary>
/// <example>
///   <summary>Extract text and Markdown from a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficePdfText -Path .\Examples\Documents\Report.pdf -PageRange '1'
/// Get-OfficePdfText -Path .\Examples\Documents\Report.pdf -AsMarkdown -OutputPath .\Examples\Documents\ReportText.md</code>
///   <para>Reads plain text directly and writes Markdown readback to a file.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfText")]
public sealed class GetOfficePdfTextCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Return logical Markdown instead of plain text.</summary>
    [Parameter]
    public SwitchParameter AsMarkdown { get; set; }

    /// <summary>Optional output text file path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path));
        var text = AsMarkdown.IsPresent
            ? string.IsNullOrWhiteSpace(PageRange) ? document.Read.Markdown() : document.Read.Markdown(PageRange!)
            : string.IsNullOrWhiteSpace(PageRange) ? document.Read.Text() : document.Read.Text(PageRange!);

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            PdfCommandUtilities.EnsureDirectory(outputPath);
            File.WriteAllText(outputPath, text);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        WriteObject(text);
    }
}
