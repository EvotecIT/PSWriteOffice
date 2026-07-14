using System;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PSWriteOffice.Services.Markdown;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Markdown;

/// <summary>Creates a Markdown document using a DSL scriptblock.</summary>
/// <para>Runs the scriptblock against a Markdown document and saves it to disk unless <c>-NoSave</c> is specified.</para>
/// <example>
///   <summary>Create a Markdown document with headings and a table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\README.md { MarkdownHeading -Level 1 -Text 'Report'; MarkdownTable -InputObject $data }</code>
///   <para>Creates a README file with a heading and table content.</para>
/// </example>
/// <example>
///   <summary>Create a report with multiple tables.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeMarkdown -Path .\Report.md {
///     MarkdownHeading -Level 1 -Text 'Summary'
///     MarkdownTable -InputObject $summary
///     MarkdownHeading -Level 2 -Text 'Details'
///     MarkdownTable -InputObject $details
///   }</code>
///   <para>Creates a report with two tables separated by headings.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeMarkdown", SupportsShouldProcess = true)]
[Alias("MarkdownNew")]
[OutputType(typeof(FileInfo), typeof(MarkdownDoc))]
public sealed class NewOfficeMarkdownCommand : PSCmdlet
    , IMarkdownWriteOptionSource
    , IMarkdownPdfOptionSource
{
    /// <summary>Destination path for the Markdown file.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "Path")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>DSL scriptblock describing Markdown content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

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
    public OfficeVisualThemeKind? PdfTheme { get; set; }

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

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fullPath = GetResolvedPath();
        if (!NoSave.IsPresent && !PdfCommandUtilities.ShouldWrite(this, fullPath, "Write new Markdown document"))
        {
            return;
        }

        var document = MarkdownDoc.Create();
        if (Content != null)
        {
            using (MarkdownDslContext.Enter(document))
            {
                Content.InvokeReturnAsIs();
            }
        }

        if (NoSave.IsPresent)
        {
            WriteObject(document);
            return;
        }

        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        File.WriteAllText(fullPath, document.ToMarkdown(MarkdownOptionUtilities.BuildWriteOptions(this)), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        SavePdfIfRequested(document, Path.GetDirectoryName(fullPath));

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(fullPath));
        }
    }

    private string GetResolvedPath()
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    private void SavePdfIfRequested(MarkdownDoc document, string? fallbackBaseDirectory)
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, pdfPath, "Write Markdown PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(pdfPath);
        var options = MarkdownOptionUtilities.BuildPdfOptions(this, this, fallbackBaseDirectory);
        var result = document.SaveAsPdf(pdfPath, options);
        MarkdownOptionUtilities.SetPdfResultVariables(this, this, result);
    }
}
