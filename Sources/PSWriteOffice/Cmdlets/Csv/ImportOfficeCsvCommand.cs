using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Imports CSV rows as PSCustomObjects, dictionaries, or a DataTable.</summary>
/// <para>Uses the CSV header to map fields to property names.</para>
/// <example>
///   <summary>Read rows as PSCustomObjects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv | Format-Table</code>
///   <para>Imports each row as a PSCustomObject.</para>
/// </example>
/// <example>
///   <summary>Read rows as dictionaries.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv -AsHashtable | ForEach-Object { $_['Name'] }</code>
///   <para>Uses hashtables for dynamic schemas or key-based access.</para>
/// </example>
/// <example>
///   <summary>Read rows as a DataTable.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Import-OfficeCsv -Path .\data.csv -AsDataTable</code>
///   <para>Emits one DataTable per input file for database and table-oriented workflows.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeCsv", DefaultParameterSetName = ParameterSetPathDelimiter)]
public sealed class ImportOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetPathDelimiter = "PathDelimiter";
    private const string ParameterSetPathCulture = "PathCulture";
    private const string ParameterSetPathDetect = "PathDetect";
    private const string ParameterSetLiteralPathDelimiter = "LiteralPathDelimiter";
    private const string ParameterSetLiteralPathCulture = "LiteralPathCulture";
    private const string ParameterSetLiteralPathDetect = "LiteralPathDetect";
    private const string ParameterSetDocument = "Document";
    private readonly CsvPowerShellRowWriter _rowWriter = new();
    private readonly List<CsvParseError> _parseErrors = new();
    private readonly CancellationTokenSource _cancellation = new();
    private bool _asDataTable;
    private bool _asHashtable;

    /// <summary>CSV document to read when already loaded.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public CsvDocument? Document { get; set; }

    /// <summary>Path to one or more CSV files. Wildcards are supported.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetPathDetect)]
    [Alias("FilePath")]
    public string[]? Path { get; set; }

    /// <summary>Literal path to one or more CSV files.</summary>
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipelineByPropertyName = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Alias("PSPath", "LP")]
    public string[]? LiteralPath { get; set; }

    /// <summary>Treat the first record as data and generate default column names.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Explicit header names to use; when provided, the first CSV record is treated as data.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public string[]? Header { get; set; }

    /// <summary>Number of parsed CSV records to skip before header discovery or data output.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [ValidateRange(0, int.MaxValue)]
    public int SkipRows { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Detect the delimiter from the first meaningful records.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathDetect)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    public SwitchParameter DetectDelimiter { get; set; }

    /// <summary>Delimiter candidates to consider when detecting the delimiter.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public bool TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Skip comment rows starting with # while discovering the header.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public bool SkipCommentRowsBeforeHeader { get; set; } = true;

    /// <summary>Skip comment rows throughout the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public SwitchParameter SkipCommentRows { get; set; }

    /// <summary>Character that identifies comment rows.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public char CommentCharacter { get; set; } = '#';

    /// <summary>Recognize W3C Extended Log File Format #Fields: rows as headers.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public bool RecognizeW3CFieldsHeader { get; set; } = true;

    /// <summary>Controls how rows with fewer or more fields than the header are handled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>Load mode controlling materialization.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public CsvLoadMode Mode { get; set; } = CsvLoadMode.Stream;

    /// <summary>Culture used for type conversions.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when reading the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public Encoding? Encoding { get; set; }

    /// <summary>Compression used when reading CSV files.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public CsvCompressionType CompressionType { get; set; } = CsvCompressionType.None;

    /// <summary>Maximum decompressed bytes allowed when reading compressed CSV files.</summary>
    [Parameter]
    public long? MaxDecompressedBytes { get; set; }

    /// <summary>How malformed quoted fields are handled.</summary>
    [Parameter]
    public CsvQuoteParsingMode QuoteParsingMode { get; set; } = CsvQuoteParsingMode.Lenient;

    /// <summary>How duplicate header names are handled.</summary>
    [Parameter]
    public CsvDuplicateHeaderBehavior DuplicateHeaderBehavior { get; set; } = CsvDuplicateHeaderBehavior.Throw;

    /// <summary>Token materialized as null when loading rows.</summary>
    [Parameter]
    public string? NullValue { get; set; }

    /// <summary>Additional date/time formats used by typed conversions.</summary>
    [Parameter]
    public string[]? DateTimeFormats { get; set; }

    /// <summary>How parse errors are handled.</summary>
    [Parameter]
    public CsvParseErrorAction ParseErrorAction { get; set; } = CsvParseErrorAction.Throw;

    /// <summary>Collect parse errors and write them as non-terminating errors after each file.</summary>
    [Parameter]
    public SwitchParameter CollectParseErrors { get; set; }

    /// <summary>Maximum number of collected parse errors before parsing fails.</summary>
    [Parameter]
    [ValidateRange(0, int.MaxValue)]
    public int MaxParseErrors { get; set; } = 100;

    /// <summary>Maximum length allowed for any parsed field.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxFieldLength { get; set; }

    /// <summary>Maximum length allowed for fields parsed from quoted records.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxQuotedFieldLength { get; set; }

    /// <summary>Normalize curly quote characters to straight quotes.</summary>
    [Parameter]
    public SwitchParameter NormalizeQuotes { get; set; }

    /// <summary>Reuse repeated string values through a per-read cache.</summary>
    [Parameter]
    public SwitchParameter InternStrings { get; set; }

    /// <summary>Report progress every N parsed records.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? ProgressInterval { get; set; }

    /// <summary>Infer DataTable column types when -AsDataTable is used.</summary>
    [Parameter]
    public SwitchParameter InferSchema { get; set; }

    /// <summary>Maximum row count inspected when schema inference is enabled.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int SchemaSampleSize { get; set; } = 1000;

    /// <summary>Emit dictionaries instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <summary>Emit one DataTable per input file instead of enumerating row objects.</summary>
    [Parameter]
    public SwitchParameter AsDataTable { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        CsvCommandValidation.EnsureHeaderOptions(NoHeader, Header);
        if (AsDataTable.IsPresent && AsHashtable.IsPresent)
        {
            throw new PSArgumentException("Specify either -AsDataTable or -AsHashtable, but not both.");
        }

        _asDataTable = AsDataTable.IsPresent;
        _asHashtable = AsHashtable.IsPresent;
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document;

        if (document == null)
        {
            ApplyCultureDelimiter();
            var options = CreateLoadOptions();

            foreach (var resolved in ResolveInputPaths())
            {
                if (!File.Exists(resolved))
                {
                    throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
                }

                if (Mode == CsvLoadMode.Stream)
                {
                    _rowWriter.Reset();
                    _parseErrors.Clear();
                    if (_asDataTable)
                    {
                        WriteDataTable(resolved, options, System.IO.Path.GetFileNameWithoutExtension(resolved));
                    }
                    else
                    {
                        CsvDocument.ReadRowsReusable(resolved, WriteRow, options);
                    }

                    WriteCollectedParseErrors(resolved);

                    continue;
                }

                _parseErrors.Clear();
                document = CsvDocument.Load(resolved, options);
                _rowWriter.Reset();
                WriteDocumentRows(document, System.IO.Path.GetFileNameWithoutExtension(resolved));
                WriteCollectedParseErrors(resolved);
            }

            return;
        }

        _rowWriter.Reset();
        WriteDocumentRows(document);
    }

    private void WriteDocumentRows(CsvDocument document, string? tableName = null)
    {
        if (_asDataTable)
        {
            WriteDataTable(document, tableName);
            return;
        }

        _rowWriter.WriteDocumentRows(document, _asHashtable, this);
    }

    private void WriteDataTable(CsvDocument document, string? tableName)
    {
        WriteObject(PSObject.AsPSObject(document.ToDataTable(CreateDataTableOptions(tableName))), enumerateCollection: false);
    }

    private void WriteDataTable(string path, CsvLoadOptions options, string? tableName)
    {
        var document = CsvDocument.Load(path, options);
        WriteObject(PSObject.AsPSObject(document.ToDataTable(CreateDataTableOptions(tableName))), enumerateCollection: false);
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        _cancellation.Cancel();
        base.StopProcessing();
    }

    private void ApplyCultureDelimiter()
    {
        if (!UseCulture.IsPresent)
        {
            return;
        }

        var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
        if (string.IsNullOrEmpty(separator) || separator.Length != 1)
        {
            throw new PSArgumentException("The selected culture must provide a single-character list separator.");
        }

        Delimiter = separator[0];
    }

    private bool IsLiteralPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetLiteralPathDetect, StringComparison.OrdinalIgnoreCase);

    private IEnumerable<string> ResolveInputPaths()
    {
        if (IsLiteralPathParameterSet())
        {
            foreach (var literalPath in LiteralPath ?? Array.Empty<string>())
            {
                yield return SessionState.Path.GetUnresolvedProviderPathFromPSPath(literalPath);
            }

            yield break;
        }

        foreach (var path in Path ?? Array.Empty<string>())
        {
            if (!WildcardPattern.ContainsWildcardCharacters(path))
            {
                yield return SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
                continue;
            }

            var resolvedPaths = SessionState.Path.GetResolvedProviderPathFromPSPath(path, out _);
            foreach (var resolvedPath in resolvedPaths)
            {
                yield return resolvedPath;
            }
        }
    }

    private CsvLoadOptions CreateLoadOptions()
    {
        var options = new CsvLoadOptions
        {
            HasHeaderRow = Header is null && !NoHeader.IsPresent,
            Header = Header,
            SkipInitialRecords = SkipRows,
            Delimiter = Delimiter,
            DetectDelimiter = DetectDelimiter.IsPresent,
            DelimiterCandidates = DelimiterCandidates,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
            SkipCommentRowsBeforeHeader = SkipCommentRowsBeforeHeader,
            SkipCommentRows = SkipCommentRows.IsPresent,
            CommentCharacter = CommentCharacter,
            RecognizeW3CFieldsHeader = RecognizeW3CFieldsHeader,
            DuplicateHeaderBehavior = DuplicateHeaderBehavior,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy,
            Mode = Mode,
            CompressionType = CompressionType,
            MaxDecompressedBytes = MaxDecompressedBytes,
            CancellationToken = _cancellation.Token,
            ProgressReportInterval = ProgressInterval ?? 0,
            ProgressCallback = ProgressInterval.HasValue ? WriteCsvProgress : null,
            QuoteParsingMode = QuoteParsingMode,
            NullValue = NullValue,
            DateTimeFormats = DateTimeFormats,
            ParseErrorAction = ParseErrorAction,
            CollectParseErrors = CollectParseErrors.IsPresent,
            MaxParseErrors = MaxParseErrors,
            ParseErrors = _parseErrors,
            MaxFieldLength = MaxFieldLength,
            MaxQuotedFieldLength = MaxQuotedFieldLength,
            NormalizeQuotes = NormalizeQuotes.IsPresent,
            InternStrings = InternStrings.IsPresent
        };

        if (Culture != null)
        {
            options.Culture = Culture;
        }

        if (Encoding != null)
        {
            options.Encoding = Encoding;
        }

        return options;
    }

    private void WriteRow(IReadOnlyList<string> header, IReadOnlyList<string> row)
    {
        _rowWriter.WriteRow(header, row, _asHashtable, this);
    }

    private CsvDataTableOptions CreateDataTableOptions(string? tableName) =>
        new()
        {
            TableName = tableName,
            InferSchema = InferSchema.IsPresent,
            SchemaSampleSize = SchemaSampleSize
        };

    private void WriteCsvProgress(CsvProgress progress)
    {
        WriteProgress(new ProgressRecord(
            activityId: 1,
            activity: "Importing CSV",
            statusDescription: $"{progress.RecordsRead:N0} records parsed"));
    }

    private void WriteCollectedParseErrors(string path)
    {
        foreach (var error in _parseErrors)
        {
            WriteError(new ErrorRecord(
                error.Exception,
                "CsvParseError",
                ErrorCategory.ParserError,
                path)
            {
                ErrorDetails = new ErrorDetails(error.Message)
            });
        }
    }
}
