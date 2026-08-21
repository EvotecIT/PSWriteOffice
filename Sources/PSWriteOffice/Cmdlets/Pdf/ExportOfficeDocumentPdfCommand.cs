using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Pdf;
using PSWriteOffice.Services.PowerPoint;
using PSWriteOffice.Services.Word;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Exports a Word, Excel, PowerPoint, Markdown, or RTF document to PDF.</summary>
/// <para>Accepts either a live OfficeIMO document from the pipeline or a supported source file.</para>
/// <example>
///   <summary>Export a live Word document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$document | Export-OfficeDocumentPdf -Path .\Report.pdf</code>
/// </example>
/// <example>
///   <summary>Export a supported file without opening it explicitly.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Export-OfficeDocumentPdf -InputPath .\Report.docx -Path .\Report.pdf -PassThru</code>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeDocumentPdf", DefaultParameterSetName = ParameterSetDocument, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeDocumentPdfCommand : PSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Live Word, Excel, PowerPoint, Markdown, or RTF document to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = ParameterSetDocument)]
    public object Document { get; set; } = null!;

    /// <summary>Source .docx, .xlsx, .pptx, .md, .markdown, or .rtf file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("SourcePath")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open an encrypted Word, Excel, or PowerPoint source file.</summary>
    [Parameter(ParameterSetName = ParameterSetPath)]
    public string? Password { get; set; }

    /// <summary>Word-specific PDF options.</summary>
    [Parameter]
    public WordPdfSaveOptions? WordOptions { get; set; }

    /// <summary>Excel-specific PDF options.</summary>
    [Parameter]
    public ExcelPdfSaveOptions? ExcelOptions { get; set; }

    /// <summary>PowerPoint-specific PDF options.</summary>
    [Parameter]
    public PowerPointPdfSaveOptions? PowerPointOptions { get; set; }

    /// <summary>Markdown-specific PDF options.</summary>
    [Parameter]
    public MarkdownPdfSaveOptions? MarkdownOptions { get; set; }

    /// <summary>RTF-specific PDF options.</summary>
    [Parameter]
    public RtfPdfSaveOptions? RtfOptions { get; set; }

    /// <summary>Open the PDF after exporting it.</summary>
    [Parameter]
    [Alias("Show")]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the saved PDF file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var outputPath = PdfCommandUtilities.ResolvePath(this, Path);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Export document to PDF")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        object document = UnwrapDocument(Document);
        IDisposable? ownedDocument = null;

        try {
            if (ParameterSetName == ParameterSetPath) {
                document = LoadDocument(out ownedDocument);
            }

            SaveDocument(document, outputPath);
        } finally {
            ownedDocument?.Dispose();
        }

        if (Open.IsPresent) {
            FileOpenService.Open(outputPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private object LoadDocument(out IDisposable? ownedDocument) {
        var sourcePath = PdfCommandUtilities.ResolveExistingFilePath(this, InputPath);
        switch (System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()) {
            case ".docx": {
                    var document = WordDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    ownedDocument = document;
                    return document;
                }
            case ".xlsx": {
                    var document = ExcelDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    ownedDocument = document;
                    return document;
                }
            case ".pptx": {
                    var document = PowerPointDocumentService.LoadPresentation(sourcePath, Password, readOnly: true);
                    ownedDocument = document;
                    return document;
                }
            case ".md":
            case ".markdown":
                ownedDocument = null;
                return MarkdownDoc.Load(sourcePath);
            case ".rtf":
                ownedDocument = null;
                return RtfDocument.Load(sourcePath);
            default:
                throw new PSArgumentException("Supported PDF source extensions are .docx, .xlsx, .pptx, .md, .markdown, and .rtf.", nameof(InputPath));
        }
    }

    private void SaveDocument(object document, string outputPath) {
        switch (document) {
            case WordDocument word:
                word.SaveAsPdf(outputPath, WordOptions ?? new WordPdfSaveOptions()).RequireSuccess();
                return;
            case ExcelDocument excel:
                excel.SaveAsPdf(outputPath, ExcelOptions ?? new ExcelPdfSaveOptions()).RequireSuccess();
                return;
            case PowerPointPresentation powerPoint:
                powerPoint.SaveAsPdf(outputPath, PowerPointOptions ?? new PowerPointPdfSaveOptions()).RequireSuccess();
                return;
            case MarkdownDoc markdown:
                markdown.SaveAsPdf(outputPath, MarkdownOptions ?? new MarkdownPdfSaveOptions()).RequireSuccess();
                return;
            case RtfDocument rtf:
                rtf.SaveAsPdf(outputPath, RtfOptions ?? new RtfPdfSaveOptions()).RequireSuccess();
                return;
            default:
                throw new PSArgumentException(
                    $"Document type '{document?.GetType().FullName ?? "<null>"}' cannot be exported to PDF. Use a WordDocument, ExcelDocument, PowerPointPresentation, MarkdownDoc, or RtfDocument.",
                    nameof(Document));
        }
    }

    private static object UnwrapDocument(object document) {
        while (document is PSObject psObject && psObject.BaseObject != document) {
            document = psObject.BaseObject;
        }

        return document;
    }
}