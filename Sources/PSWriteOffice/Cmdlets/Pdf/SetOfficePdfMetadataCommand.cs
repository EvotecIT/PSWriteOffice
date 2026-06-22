using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets PDF document metadata on generated documents or existing PDF files.</summary>
/// <remarks>
/// In a <c>New-OfficePdf</c> script block this command updates the generated document metadata.
/// With <c>-Path</c> and <c>-OutputPath</c>, it rewrites an existing PDF with updated metadata unless <c>-Incremental</c> is used.
/// </remarks>
/// <example>
///   <summary>Set metadata while generating a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Report.pdf {
///     PdfMetadata -Title 'Service Review' -Author 'PSWriteOffice' -Subject 'Operations'
///     PdfHeading 'Service Review'
///   }</code>
///   <para>Stores metadata on a newly generated PDF.</para>
/// </example>
/// <example>
///   <summary>Rewrite metadata on an existing PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Set-OfficePdfMetadata -Path .\Input.pdf -OutputPath .\Output.pdf -Title 'Reviewed package' -Author 'Operations'</code>
///   <para>Writes a new PDF with updated metadata.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfMetadata", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfMetadata")]
[OutputType(typeof(PdfDocument), typeof(FileInfo))]
public sealed class SetOfficePdfMetadataCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";
    private const string ParameterSetFile = "File";

    /// <summary>PDF document to update outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Existing PDF path to rewrite with updated metadata.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetFile)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Output PDF path when rewriting an existing PDF.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetFile)]
    public string? OutputPath { get; set; }

    /// <summary>Document title.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Document author.</summary>
    [Parameter]
    public string? Author { get; set; }

    /// <summary>Document subject.</summary>
    [Parameter]
    public string? Subject { get; set; }

    /// <summary>Document keywords.</summary>
    [Parameter]
    public string? Keywords { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Append a metadata-only incremental PDF revision instead of rewriting the existing PDF bytes.</summary>
    [Parameter(ParameterSetName = ParameterSetFile)]
    public SwitchParameter Incremental { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetFile)
        {
            var inputPath = PdfCommandUtilities.ResolvePath(this, Path!);
            var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath!);
            PdfCommandUtilities.EnsureDirectory(outputPath);
            if (Incremental.IsPresent)
            {
                PdfIncrementalUpdater.UpdateMetadata(inputPath, outputPath, Title, Author, Subject, Keywords);
            }
            else
            {
                PdfMetadataEditor.UpdateMetadata(inputPath, outputPath, Title, Author, Subject, Keywords);
            }

            WriteObject(new FileInfo(outputPath));
            return;
        }

        var document = PdfCommandUtilities.ResolveDocument(this, Document, ParameterSetName, ParameterSetDocument);
        document.Meta(Title, Author, Subject, Keywords);
        if (PassThru.IsPresent)
        {
            WriteObject(document);
        }
    }
}
