using System.Management.Automation;
using OfficeIMO.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Reads a supported file into the OfficeIMO shared document read result envelope.</summary>
/// <remarks>
/// Use <c>-AsJson</c> for a deterministic serialized payload suitable for indexing or API handoff.
/// </remarks>
[Cmdlet(VerbsCommon.Get, "OfficeDocument")]
[Alias("Read-OfficeDocument")]
[OutputType(typeof(OfficeDocumentReadResult), typeof(string))]
public sealed class GetOfficeDocumentCommand : PSCmdlet
{
    /// <summary>File path to read.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Return the read result as JSON.</summary>
    [Parameter]
    public SwitchParameter AsJson { get; set; }

    /// <summary>Indent JSON output.</summary>
    [Parameter]
    public SwitchParameter Indented { get; set; }

    /// <summary>Maximum input size in bytes.</summary>
    [Parameter]
    public long? MaxInputBytes { get; set; }

    /// <summary>OpenXML maximum characters per part.</summary>
    [Parameter]
    public long? OpenXmlMaxCharactersInPart { get; set; }

    /// <summary>Maximum emitted chunk characters.</summary>
    [Parameter]
    public int? MaxChars { get; set; }

    /// <summary>Maximum table rows per emitted table chunk.</summary>
    [Parameter]
    public int? MaxTableRows { get; set; }

    /// <summary>Exclude Word footnotes.</summary>
    [Parameter]
    public SwitchParameter ExcludeWordFootnotes { get; set; }

    /// <summary>Exclude PowerPoint speaker notes.</summary>
    [Parameter]
    public SwitchParameter ExcludePowerPointNotes { get; set; }

    /// <summary>Treat the first Excel row as data instead of headers.</summary>
    [Parameter]
    public SwitchParameter NoExcelHeaders { get; set; }

    /// <summary>Excel rows per emitted worksheet chunk.</summary>
    [Parameter]
    public int? ExcelChunkRows { get; set; }

    /// <summary>Optional Excel sheet name to read.</summary>
    [Parameter]
    public string? ExcelSheetName { get; set; }

    /// <summary>Optional Excel A1 range to read.</summary>
    [Parameter]
    public string? ExcelA1Range { get; set; }

    /// <summary>Do not split Markdown by headings.</summary>
    [Parameter]
    public SwitchParameter NoMarkdownHeadingChunks { get; set; }

    /// <summary>Disable source and chunk hash computation.</summary>
    [Parameter]
    public SwitchParameter NoHashes { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        ReaderCommandUtilities.RegisterReaderAdapters();
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var path = ReaderCommandUtilities.ResolvePath(this, Path);
        var options = ReaderCommandUtilities.BuildReaderOptions(
            MaxInputBytes,
            OpenXmlMaxCharactersInPart,
            MaxChars,
            MaxTableRows,
            !ExcludeWordFootnotes.IsPresent,
            !ExcludePowerPointNotes.IsPresent,
            !NoExcelHeaders.IsPresent,
            ExcelChunkRows,
            ExcelSheetName,
            ExcelA1Range,
            !NoMarkdownHeadingChunks.IsPresent,
            !NoHashes.IsPresent);

        WriteObject(AsJson.IsPresent
            ? DocumentReader.ReadDocumentJson(path, options, Indented.IsPresent)
            : DocumentReader.ReadDocument(path, options));
    }
}
