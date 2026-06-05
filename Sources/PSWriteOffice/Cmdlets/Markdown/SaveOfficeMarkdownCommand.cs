using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Saves a Markdown document and optionally creates a PDF sidecar.</summary>
/// <example>
///   <summary>Save Markdown and PDF outputs.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc | Save-OfficeMarkdown -Path .\Report.md -PdfPath .\Report.pdf</code>
///   <para>Writes both artifacts from the same Markdown document model.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeMarkdown")]
[OutputType(typeof(MarkdownDoc), typeof(FileInfo))]
public sealed class SaveOfficeMarkdownCommand : PSCmdlet
{
    /// <summary>Markdown document to save.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public MarkdownDoc Document { get; set; } = null!;

    /// <summary>Destination Markdown path.</summary>
    [Parameter(Position = 1)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Optional PDF path to create from the same Markdown document.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

    /// <summary>Emit the Markdown document rather than the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        FileInfo? savedFile = null;
        if (!string.IsNullOrWhiteSpace(Path))
        {
            var fullPath = PdfCommandUtilities.ResolvePath(this, Path!);
            PdfCommandUtilities.EnsureDirectory(fullPath);
            File.WriteAllText(fullPath, Document.ToMarkdown(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            savedFile = new FileInfo(fullPath);
        }

        if (!string.IsNullOrWhiteSpace(PdfPath))
        {
            Document.SaveAsPdf(PdfCommandUtilities.ResolvePath(this, PdfPath!));
        }

        WriteObject(PassThru.IsPresent ? Document : savedFile ?? (object)Document);
    }
}
