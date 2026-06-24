using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Saves a Markdown document and optionally creates a PDF sidecar.</summary>
/// <example>
///   <summary>Save Markdown and PDF outputs.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$doc | Save-OfficeMarkdown -Path .\Report.md -PdfPath .\Report.pdf</code>
///   <para>Writes both artifacts from the same Markdown document model.</para>
/// </example>
[Cmdlet(VerbsData.Save, "OfficeMarkdown", SupportsShouldProcess = true)]
[OutputType(typeof(MarkdownDoc), typeof(FileInfo))]
public sealed class SaveOfficeMarkdownCommand : PSCmdlet
    , IMarkdownWriteOptionSource
    , IMarkdownPdfOptionSource
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

    /// <summary>Optional Markdown writer options.</summary>
    [Parameter]
    public MarkdownWriteOptions? WriteOptions { get; set; }

    /// <summary>Friendly Markdown writer profile.</summary>
    [Parameter]
    public OfficeMarkdownWriteProfile? WriteProfile { get; set; }

    /// <summary>Controls how Markdown images are serialized.</summary>
    [Parameter]
    public MarkdownImageRenderingMode? ImageRenderingMode { get; set; }

    /// <summary>Markdown line ending: CRLF, LF, CR, or a literal line ending string.</summary>
    [Parameter]
    public string? LineEnding { get; set; }

    /// <summary>Unordered list marker: '-', '*', or '+'.</summary>
    [Parameter]
    public string? UnorderedListMarker { get; set; }

    /// <summary>Advanced Markdown PDF options. Friendly PDF parameters override matching values.</summary>
    [Parameter]
    public MarkdownPdfSaveOptions? MarkdownPdfOptions { get; set; }

    /// <summary>Underlying OfficeIMO.Pdf options used by Markdown PDF export.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Built-in Markdown PDF visual theme.</summary>
    [Parameter]
    public MarkdownPdfThemeKind? PdfTheme { get; set; }

    /// <summary>Default font family used by Markdown PDF export.</summary>
    [Parameter]
    public string? PdfFontFamily { get; set; }

    /// <summary>PDF title metadata.</summary>
    [Parameter]
    public string? PdfTitle { get; set; }

    /// <summary>PDF author metadata.</summary>
    [Parameter]
    public string? PdfAuthor { get; set; }

    /// <summary>PDF subject metadata.</summary>
    [Parameter]
    public string? PdfSubject { get; set; }

    /// <summary>PDF keywords metadata.</summary>
    [Parameter]
    public string? PdfKeywords { get; set; }

    /// <summary>Base directory used to resolve local Markdown images during PDF export.</summary>
    [Parameter]
    public string? PdfBaseDirectory { get; set; }

    /// <summary>Apply the built-in Word-like Markdown PDF baseline theme.</summary>
    [Parameter]
    public bool? PdfApplyWordLikeTheme { get; set; }

    /// <summary>Embed supported local image files in Markdown PDF output.</summary>
    [Parameter]
    public bool? PdfIncludeLocalImages { get; set; }

    /// <summary>Embed supported data URI images in Markdown PDF output.</summary>
    [Parameter]
    public bool? PdfIncludeDataUriImages { get; set; }

    /// <summary>Require local images to resolve under the base directory.</summary>
    [Parameter]
    public bool? PdfRestrictLocalImagesToBaseDirectory { get; set; }

    /// <summary>Maximum decoded bytes for one data URI image in Markdown PDF output.</summary>
    [Parameter]
    public int? PdfMaximumDataUriImageBytes { get; set; }

    /// <summary>Fallback PDF image width in points.</summary>
    [Parameter]
    public double? PdfDefaultImageWidth { get; set; }

    /// <summary>Fallback PDF image height in points.</summary>
    [Parameter]
    public double? PdfDefaultImageHeight { get; set; }

    /// <summary>Controls how YAML front matter appears in the PDF body.</summary>
    [Parameter]
    public MarkdownPdfFrontMatterRenderMode? PdfFrontMatterRenderMode { get; set; }

    /// <summary>Use front matter values to select a visual theme.</summary>
    [Parameter]
    public bool? PdfUseFrontMatterVisualTheme { get; set; }

    /// <summary>Use front matter values as PDF metadata.</summary>
    [Parameter]
    public bool? PdfUseFrontMatterMetadata { get; set; }

    /// <summary>Use the first Markdown heading as the PDF title when no title is supplied.</summary>
    [Parameter]
    public bool? PdfUseFirstHeadingAsTitle { get; set; }

    /// <summary>Create PDF outlines from Markdown headings.</summary>
    [Parameter]
    public bool? PdfCreateOutlineFromHeadings { get; set; }

    /// <summary>Variable name that receives Markdown PDF export warnings.</summary>
    [Parameter]
    public string? PdfWarningVariable { get; set; }

    /// <summary>Variable name that receives the Markdown PDF conversion report.</summary>
    [Parameter]
    public string? PdfConversionReportVariable { get; set; }

    /// <summary>Emit the Markdown document rather than the saved file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (string.IsNullOrWhiteSpace(Path) && string.IsNullOrWhiteSpace(PdfPath) && !PassThru.IsPresent)
        {
            throw new PSInvalidOperationException("Use -Path, -PdfPath, or -PassThru when saving a Markdown document.");
        }

        FileInfo? savedFile = null;
        if (!string.IsNullOrWhiteSpace(Path))
        {
            var fullPath = PdfCommandUtilities.ResolvePath(this, Path!);
            if (!PdfCommandUtilities.ShouldWrite(this, fullPath, "Save Markdown document"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(fullPath);
            File.WriteAllText(fullPath, Document.ToMarkdown(MarkdownOptionUtilities.BuildWriteOptions(this)), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            savedFile = new FileInfo(fullPath);
        }

        if (!string.IsNullOrWhiteSpace(PdfPath))
        {
            var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
            if (!PdfCommandUtilities.ShouldWrite(this, pdfPath, "Write Markdown PDF"))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(pdfPath);
            var options = MarkdownOptionUtilities.BuildPdfOptions(this, this, ResolvePdfBaseDirectory(savedFile));
            Document.SaveAsPdf(pdfPath, options);
            MarkdownOptionUtilities.SetPdfResultVariables(this, this, options);
        }

        WriteObject(PassThru.IsPresent ? Document : savedFile ?? (object)Document);
    }

    private string? ResolvePdfBaseDirectory(FileInfo? savedFile)
    {
        if (savedFile?.DirectoryName != null)
        {
            return savedFile.DirectoryName;
        }

        if (!string.IsNullOrWhiteSpace(PdfPath))
        {
            var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
            return System.IO.Path.GetDirectoryName(pdfPath);
        }

        return null;
    }
}
