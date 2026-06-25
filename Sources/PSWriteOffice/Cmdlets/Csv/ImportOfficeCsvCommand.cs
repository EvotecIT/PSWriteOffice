using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Imports CSV rows as PSCustomObjects or dictionaries.</summary>
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
///   <summary>Convert CSV text into rows.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = ConvertFrom-OfficeCsv -Text "Name,Value`nAlpha,1"</code>
///   <para>Parses CSV text and emits rows without writing a temporary file.</para>
/// </example>
[Cmdlet(VerbsData.Import, "OfficeCsv", DefaultParameterSetName = ParameterSetPathDelimiter)]
[Alias("Get-OfficeCsvData", "ConvertFrom-OfficeCsv")]
public sealed class ImportOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetPathDelimiter = "PathDelimiter";
    private const string ParameterSetPathCulture = "PathCulture";
    private const string ParameterSetPathDetect = "PathDetect";
    private const string ParameterSetLiteralPathDelimiter = "LiteralPathDelimiter";
    private const string ParameterSetLiteralPathCulture = "LiteralPathCulture";
    private const string ParameterSetLiteralPathDetect = "LiteralPathDetect";
    private const string ParameterSetTextDelimiter = "TextDelimiter";
    private const string ParameterSetTextCulture = "TextCulture";
    private const string ParameterSetTextDetect = "TextDetect";
    private const string ParameterSetDocument = "Document";
    private readonly StringBuilder _textInput = new();
    private bool _asHashtable;
    private bool _hasTextInput;
    private bool _prevalidatedOutputProperties;

    /// <summary>CSV document to read when already loaded.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public CsvDocument? Document { get; set; }

    /// <summary>Path to a CSV file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPathDetect)]
    [Alias("FilePath")]
    public string? Path { get; set; }

    /// <summary>Literal path to a CSV file.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Alias("PSPath", "LP")]
    public string? LiteralPath { get; set; }

    /// <summary>CSV text to parse.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetTextCulture)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetTextDetect)]
    public string? Text { get; set; }

    /// <summary>Treat the first record as data and generate default column names.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Explicit header names to use; when provided, the first CSV record is treated as data.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public string[]? Header { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Detect the delimiter from the first meaningful records.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathDetect)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter DetectDelimiter { get; set; }

    /// <summary>Delimiter candidates to consider when detecting the delimiter.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public char[]? DelimiterCandidates { get; set; }

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetTextCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public bool TrimWhitespace { get; set; }

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Skip comment rows starting with # while discovering the header.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public bool SkipCommentRowsBeforeHeader { get; set; } = true;

    /// <summary>Skip comment rows throughout the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public SwitchParameter SkipCommentRows { get; set; }

    /// <summary>Character that identifies comment rows.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public char CommentCharacter { get; set; } = '#';

    /// <summary>Recognize W3C Extended Log File Format #Fields: rows as headers.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public bool RecognizeW3CFieldsHeader { get; set; } = true;

    /// <summary>Controls how rows with fewer or more fields than the header are handled.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvColumnCountMismatchPolicy ColumnCountMismatchPolicy { get; set; } = CsvColumnCountMismatchPolicy.PadMissingFieldsAndIgnoreExtraFields;

    /// <summary>Load mode controlling materialization.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CsvLoadMode Mode { get; set; } = CsvLoadMode.Stream;

    /// <summary>Culture used for type conversions.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    [Parameter(ParameterSetName = ParameterSetTextDelimiter)]
    [Parameter(ParameterSetName = ParameterSetTextCulture)]
    [Parameter(ParameterSetName = ParameterSetTextDetect)]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when reading the file.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public Encoding? Encoding { get; set; }

    /// <summary>Emit dictionaries instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        _asHashtable = AsHashtable.IsPresent;
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (IsTextParameterSet())
        {
            if (_hasTextInput)
            {
                _textInput.AppendLine();
            }

            _textInput.Append(Text ?? string.Empty);
            _hasTextInput = true;
            return;
        }

        var document = Document;

        if (document == null)
        {
            ApplyCultureDelimiter();
            var options = CreateLoadOptions();

            var unresolvedPath = IsLiteralPathParameterSet() ? LiteralPath : Path;
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(unresolvedPath!);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }

            if (Mode == CsvLoadMode.Stream)
            {
                CsvDocument.ReadRowsReusable(resolved, WriteRow, options);
                return;
            }

            document = CsvDocument.Load(resolved, options);
        }

        WriteDocumentRows(document);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (!IsTextParameterSet())
        {
            return;
        }

        ApplyCultureDelimiter();
        var options = CreateLoadOptions();
        var csvText = _textInput.ToString();

        if (Mode == CsvLoadMode.Stream)
        {
            using var reader = new StringReader(csvText);
            CsvDocument.ReadRowsReusable(reader, WriteRow, options);
            return;
        }

        WriteDocumentRows(CsvDocument.Parse(csvText, options));
    }

    private void WriteDocumentRows(CsvDocument document)
    {
        var header = document.Header;
        foreach (var row in document.AsEnumerable())
        {
            if (AsHashtable.IsPresent)
            {
                var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                for (var i = 0; i < header.Count && i < row.FieldCount; i++)
                {
                    rowValues[header[i]] = row[i];
                }

                WriteObject(rowValues);
            }
            else
            {
                var psObj = new PSObject(header.Count);
                for (var i = 0; i < header.Count && i < row.FieldCount; i++)
                {
                    psObj.Properties.Add(new PSNoteProperty(header[i], row[i]), _prevalidatedOutputProperties);
                }

                _prevalidatedOutputProperties = true;
                WriteObject(psObj);
            }
        }
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

    private bool IsTextParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetTextDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetTextCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetTextDetect, StringComparison.OrdinalIgnoreCase);

    private CsvLoadOptions CreateLoadOptions()
    {
        var options = new CsvLoadOptions
        {
            HasHeaderRow = Header is null && !NoHeader.IsPresent,
            Header = Header,
            Delimiter = Delimiter,
            DetectDelimiter = DetectDelimiter.IsPresent,
            DelimiterCandidates = DelimiterCandidates,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
            SkipCommentRowsBeforeHeader = SkipCommentRowsBeforeHeader,
            SkipCommentRows = SkipCommentRows.IsPresent,
            CommentCharacter = CommentCharacter,
            RecognizeW3CFieldsHeader = RecognizeW3CFieldsHeader,
            ColumnCountMismatchPolicy = ColumnCountMismatchPolicy,
            Mode = Mode
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
        var headerCount = header.Count;
        var rowCount = row.Count;
        var valueCount = rowCount < headerCount ? rowCount : headerCount;

        if (_asHashtable)
        {
            var rowValues = new Dictionary<string, object?>(valueCount, StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < valueCount; i++)
            {
                rowValues[header[i]] = row[i];
            }

            WriteObject(rowValues);
            return;
        }

        var psObj = new PSObject(headerCount);
        var prevalidated = _prevalidatedOutputProperties;
        for (var i = 0; i < valueCount; i++)
        {
            psObj.Properties.Add(new PSNoteProperty(header[i], row[i]), prevalidated);
        }

        _prevalidatedOutputProperties = true;
        WriteObject(psObj);
    }
}
