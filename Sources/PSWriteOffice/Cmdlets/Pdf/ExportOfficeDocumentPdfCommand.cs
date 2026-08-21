using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
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
/// <example>
///   <summary>Configure Markdown PDF export with ordinary PowerShell parameters.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeMarkdownPdfOptions -Title 'Service report' -IncludeLocalImages -BaseDirectory .\Assets
/// Export-OfficeDocumentPdf -InputPath .\Report.md -Path .\Report.pdf -MarkdownOptions $options</code>
///   <para>The New-Office*PdfOptions commands build every format-specific options object; no hashtable or .NET constructor is required.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeDocumentPdf", DefaultParameterSetName = ParameterSetDocument, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeDocumentPdfCommand : PSCmdlet {
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetPath = "Path";

    /// <summary>Live Word, Excel, PowerPoint, Markdown, or RTF document to export. Saved FileInfo and path strings from the pipeline are opened automatically.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0, ParameterSetName = ParameterSetDocument)]
    public object Document { get; set; } = null!;

    /// <summary>Source .docx, .xlsx, .pptx, .md, .markdown, or .rtf file.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("SourcePath", "FullName")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Destination PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutputPath", "FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Password used to open an encrypted Word, Excel, or PowerPoint source file.</summary>
    [Parameter]
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

    /// <summary>Variable name that receives structured PDF conversion warnings.</summary>
    [Parameter]
    public string? PdfWarningVariable { get; set; }

    /// <summary>Variable name that receives the structured PDF conversion report.</summary>
    [Parameter]
    public string? PdfConversionReportVariable { get; set; }

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
        object document;
        Action? closeOwnedDocument = null;
        string? sourcePath = null;

        try {
            if (ParameterSetName == ParameterSetPath) {
                document = LoadDocument(InputPath, out closeOwnedDocument, out sourcePath);
            } else {
                document = UnwrapDocument(Document);
                if (document is FileInfo file) {
                    document = LoadDocument(file.FullName, out closeOwnedDocument, out sourcePath);
                } else if (document is string path) {
                    document = LoadDocument(path, out closeOwnedDocument, out sourcePath);
                }
            }

            PdfSaveResult result = SaveDocument(document, outputPath, sourcePath);
            PdfCommandUtilities.SetVariable(this, PdfWarningVariable, result.Warnings);
            PdfCommandUtilities.SetVariable(this, PdfConversionReportVariable, result.Report);
            result.RequireSuccess();
        } finally {
            closeOwnedDocument?.Invoke();
        }

        if (Open.IsPresent) {
            FileOpenService.Open(outputPath);
        }

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(outputPath));
        }
    }

    private object LoadDocument(string inputPath, out Action? closeOwnedDocument, out string sourcePath) {
        sourcePath = PdfCommandUtilities.ResolveExistingFilePath(this, inputPath);
        switch (System.IO.Path.GetExtension(sourcePath).ToLowerInvariant()) {
            case ".docx": {
                    var document = WordDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    closeOwnedDocument = () => WordDocumentService.CloseDocument(document);
                    return document;
                }
            case ".xlsx": {
                    var document = ExcelDocumentService.LoadDocument(sourcePath, readOnly: true, autoSave: false, Password);
                    closeOwnedDocument = () => ExcelDocumentService.CloseDocument(document);
                    return document;
                }
            case ".pptx": {
                    var document = PowerPointDocumentService.LoadPresentation(sourcePath, Password, readOnly: true);
                    closeOwnedDocument = () => PowerPointDocumentService.ClosePresentation(document, save: false, show: false);
                    return document;
                }
            case ".md":
            case ".markdown":
                closeOwnedDocument = null;
                return MarkdownDoc.Load(sourcePath);
            case ".rtf":
                closeOwnedDocument = null;
                return RtfDocument.Load(sourcePath);
            default:
                throw new PSArgumentException("Supported PDF source extensions are .docx, .xlsx, .pptx, .md, .markdown, and .rtf.", nameof(InputPath));
        }
    }

    private PdfSaveResult SaveDocument(object document, string outputPath, string? sourcePath) {
        switch (document) {
            case WordDocument word:
                return word.SaveAsPdf(outputPath, WordOptions ?? new WordPdfSaveOptions());
            case ExcelDocument excel:
                return excel.SaveAsPdf(outputPath, ExcelOptions ?? new ExcelPdfSaveOptions());
            case PowerPointPresentation powerPoint:
                return powerPoint.SaveAsPdf(outputPath, PowerPointOptions ?? new PowerPointPdfSaveOptions());
            case MarkdownDoc markdown:
                return markdown.SaveAsPdf(outputPath, PrepareMarkdownOptions(sourcePath));
            case RtfDocument rtf:
                return rtf.SaveAsPdf(outputPath, RtfOptions ?? new RtfPdfSaveOptions());
            default:
                throw new PSArgumentException(
                    $"Document type '{document?.GetType().FullName ?? "<null>"}' cannot be exported to PDF. Use a WordDocument, ExcelDocument, PowerPointPresentation, MarkdownDoc, or RtfDocument.",
                    nameof(Document));
        }
    }

    private MarkdownPdfSaveOptions PrepareMarkdownOptions(string? sourcePath) {
        var options = MarkdownOptions?.Clone() ?? new MarkdownPdfSaveOptions();
        if (options.ResourcePolicy.AllowLocalFileAccess &&
            string.IsNullOrWhiteSpace(options.BaseDirectory) &&
            !string.IsNullOrWhiteSpace(sourcePath)) {
            options.BaseDirectory = System.IO.Path.GetDirectoryName(sourcePath);
        }

        return options;
    }

    private static object UnwrapDocument(object document) {
        while (document is PSObject psObject && psObject.BaseObject != document) {
            document = psObject.BaseObject;
        }

        return document;
    }
}
