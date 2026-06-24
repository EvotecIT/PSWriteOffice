using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Reads CSV rows as PSCustomObjects or dictionaries.</summary>
/// <para>Uses the CSV header to map fields to property names.</para>
/// <example>
///   <summary>Read rows as PSCustomObjects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeCsvData -Path .\data.csv | Format-Table</code>
///   <para>Returns each row as a PSCustomObject.</para>
/// </example>
/// <example>
///   <summary>Read rows as dictionaries.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeCsvData -Path .\data.csv -AsHashtable | ForEach-Object { $_['Name'] }</code>
///   <para>Uses hashtables for dynamic schemas or key-based access.</para>
/// </example>
/// <example>
///   <summary>Read semicolon-delimited data without headers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Get-OfficeCsvData -Path .\data.csv -Delimiter ';' -NoHeader</code>
///   <para>Reads CSV files that lack a header row and use a custom delimiter.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeCsvData", DefaultParameterSetName = ParameterSetPathDelimiter)]
public sealed class GetOfficeCsvDataCommand : PSCmdlet
{
    private const string ParameterSetPathDelimiter = "PathDelimiter";
    private const string ParameterSetPathCulture = "PathCulture";
    private const string ParameterSetPathDetect = "PathDetect";
    private const string ParameterSetLiteralPathDelimiter = "LiteralPathDelimiter";
    private const string ParameterSetLiteralPathCulture = "LiteralPathCulture";
    private const string ParameterSetLiteralPathDetect = "LiteralPathDetect";
    private const string ParameterSetDocument = "Document";
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

    /// <summary>Indicates whether the first record is a header row.</summary>
    [Parameter(ParameterSetName = ParameterSetPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetPathCulture)]
    [Parameter(ParameterSetName = ParameterSetPathDetect)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathCulture)]
    [Parameter(ParameterSetName = ParameterSetLiteralPathDetect)]
    public bool HasHeaderRow { get; set; } = true;

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

    /// <summary>Emit dictionaries instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document;

        if (document == null)
        {
            ApplyCultureDelimiter();

            var options = new CsvLoadOptions
            {
                HasHeaderRow = Header is null && !NoHeader.IsPresent && HasHeaderRow,
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

            var unresolvedPath = IsLiteralPathParameterSet() ? LiteralPath : Path;
            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(unresolvedPath!);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }

            if (Mode == CsvLoadMode.Stream)
            {
                CsvDocument.ReadRows(resolved, WriteRow, options);
                return;
            }

            document = CsvDocument.Load(resolved, options);
        }

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

    private void WriteRow(IReadOnlyList<string> header, IReadOnlyList<string> row)
    {
        if (AsHashtable.IsPresent)
        {
            var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < header.Count && i < row.Count; i++)
            {
                rowValues[header[i]] = row[i];
            }

            WriteObject(rowValues);
            return;
        }

        var psObj = new PSObject(header.Count);
        for (var i = 0; i < header.Count && i < row.Count; i++)
        {
            psObj.Properties.Add(new PSNoteProperty(header[i], row[i]), _prevalidatedOutputProperties);
        }

        _prevalidatedOutputProperties = true;
        WriteObject(psObj);
    }
}
