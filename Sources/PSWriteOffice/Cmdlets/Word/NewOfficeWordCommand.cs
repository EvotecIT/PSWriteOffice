using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Word;

/// <summary>Creates a Word document using the DSL.</summary>
/// <para>Handles file creation or template cloning, scriptblock execution, optional autosave, and emits the document path when <c>-PassThru</c> is used.</para>
/// <example>
///   <summary>Create a document inline.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -Path .\Report.docx { WordSection { WordParagraph 'Hello DSL' } } -Open</code>
///   <para>Builds a document, adds one paragraph, saves it to disk, and opens it.</para>
/// </example>
/// <example>
///   <summary>Create a document from a template.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeWord -TemplatePath .\Template.docx -Path .\Report.docx { WordParagraph -Text 'Generated content' -StyleId 'ReportBody' }</code>
///   <para>Copies the template to the output path, runs the DSL against the copied document, and saves it.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeWord", SupportsShouldProcess = true)]
[Alias("WordNew")]
public sealed class NewOfficeWordCommand : PSCmdlet
{
    /// <summary>Destination path for the document.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "Path")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Existing .docx file to clone before running the DSL.</summary>
    [Parameter]
    public string? TemplatePath { get; set; }

    /// <summary>DSL scriptblock describing document content.</summary>
    [Parameter(Position = 1)]
    public ScriptBlock? Content { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for chaining.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Open the document after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Skip saving after executing the DSL.</summary>
    [Parameter]
    public SwitchParameter NoSave { get; set; }

    /// <summary>Enable OfficeIMO AutoSave mode.</summary>
    [Parameter]
    public SwitchParameter AutoSave { get; set; }

    /// <summary>Password used to save the document as an encrypted package.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>Optional PDF path to create from the same Word document before closing it.</summary>
    [Parameter]
    public string? PdfPath { get; set; }

    /// <summary>Optional default font family used by the native Word PDF converter.</summary>
    [Parameter]
    public string? PdfFontFamily { get; set; }

    /// <summary>Allow the native Word PDF converter to embed installed system fonts used by the document.</summary>
    [Parameter]
    [Alias("AllowSystemFontEmbedding")]
    public SwitchParameter PdfAllowSystemFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (!NoSave.IsPresent && AutoSave.IsPresent && !string.IsNullOrEmpty(Password))
        {
            throw new PSArgumentException("Encrypted Word documents require explicit Save-OfficeWord -Password or Close-OfficeWord -Save -Password. -AutoSave cannot be used with -Password.");
        }

        var fullPath = GetResolvedPath();
        var action = NoSave.IsPresent
            ? string.IsNullOrWhiteSpace(TemplatePath)
                ? "Create in-memory Word document"
                : "Create Word document from template"
            : "Write new Word document";
        if (!PdfCommandUtilities.ShouldWrite(this, fullPath, action))
        {
            return;
        }

        if (!NoSave.IsPresent || !string.IsNullOrWhiteSpace(TemplatePath))
        {
            var directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
        }

        var document = CreateOrLoadDocument(fullPath);

        if (Content == null)
        {
            WriteObject(document);
            return;
        }

        WordDocumentService.InvokeDsl(document, Content);

        if (NoSave.IsPresent)
        {
            WriteObject(document);
            return;
        }

        SavePdfIfRequested(document);
        WordDocumentService.SaveDocument(document, Open.IsPresent, fullPath, Password);

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

    private OfficeIMO.Word.WordDocument CreateOrLoadDocument(string fullPath)
    {
        if (string.IsNullOrWhiteSpace(TemplatePath))
        {
            if (NoSave.IsPresent)
            {
                return WordDocumentService.CreateInMemoryDocument();
            }

            return WordDocumentService.CreateDocument(fullPath, AutoSave.IsPresent);
        }

        var templatePath = ResolveFileSystemPath(TemplatePath!);
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException($"Template file {templatePath} doesn't exist.", templatePath);
        }

        if (!string.Equals(templatePath, fullPath, StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(templatePath, fullPath, overwrite: true);
        }

        return WordDocumentService.LoadDocument(fullPath, readOnly: false, autoSave: AutoSave.IsPresent);
    }

    private string ResolveFileSystemPath(string path)
    {
        var providerPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
        return Path.IsPathRooted(providerPath)
            ? providerPath
            : Path.Combine(SessionState.Path.CurrentFileSystemLocation.Path, providerPath);
    }

    private void SavePdfIfRequested(OfficeIMO.Word.WordDocument document)
    {
        if (string.IsNullOrWhiteSpace(PdfPath))
        {
            return;
        }

        var pdfPath = PdfCommandUtilities.ResolvePath(this, PdfPath!);
        if (!PdfCommandUtilities.ShouldWrite(this, pdfPath, "Write Word PDF"))
        {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(pdfPath);
        if (PdfAllowSystemFontEmbedding.IsPresent || !string.IsNullOrWhiteSpace(PdfFontFamily))
        {
            var pdfOptions = new PdfSaveOptions
            {
                FontFamily = PdfFontFamily
            };
            pdfOptions.ResourcePolicy.AllowSystemFontEmbedding = PdfAllowSystemFontEmbedding.IsPresent;
            document.SaveAsPdf(pdfPath, pdfOptions).RequireSuccess();
            return;
        }

        document.SaveAsPdf(pdfPath).RequireSuccess();
    }
}
