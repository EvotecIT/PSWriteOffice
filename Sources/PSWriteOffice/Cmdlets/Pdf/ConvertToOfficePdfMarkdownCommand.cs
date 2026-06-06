using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts PDF logical text readback to Markdown.</summary>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfMarkdown")]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficePdfMarkdownCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Optional output Markdown file path.</summary>
    [Parameter]
    public string? OutputPath { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = PdfDocument.Open(PdfCommandUtilities.ResolvePath(this, Path));
        var markdown = string.IsNullOrWhiteSpace(PageRange)
            ? document.Read.Markdown()
            : document.Read.Markdown(PageRange!);

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            PdfCommandUtilities.EnsureDirectory(outputPath);
            File.WriteAllText(outputPath, markdown);
            WriteObject(new FileInfo(outputPath));
            return;
        }

        WriteObject(markdown);
    }
}
