using System;
using System.IO;
using System.Management.Automation;
using System.Text;
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
[Cmdlet(VerbsCommon.New, "OfficeMarkdown")]
[OutputType(typeof(FileInfo), typeof(MarkdownDoc))]
public sealed class NewOfficeMarkdownCommand : PSCmdlet
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

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var fullPath = GetResolvedPath();
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
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

        File.WriteAllText(fullPath, document.ToMarkdown(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        SavePdfIfRequested(document);

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

    private void SavePdfIfRequested(MarkdownDoc document)
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
        PdfCommandUtilities.EnsureDirectory(pdfPath);
        document.SaveAsPdf(pdfPath);
    }
}
