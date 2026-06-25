using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Exports objects or a CSV document to a CSV file.</summary>
/// <para>Use <c>ConvertTo-OfficeCsv</c> when the destination should be CSV text in the pipeline.</para>
/// <example>
///   <summary>Export objects to a CSV file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data | Export-OfficeCsv -Path .\export.csv</code>
///   <para>Streams PowerShell objects into a CSV file.</para>
/// </example>
/// <example>
///   <summary>Export with culture list separator.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data | Export-OfficeCsv -Path .\export.csv -UseCulture -Culture pl-PL</code>
///   <para>Uses the selected culture list separator as the delimiter.</para>
/// </example>
[Cmdlet(VerbsData.Export, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObjectPathDelimiter, SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ExportOfficeCsvCommand : PSCmdlet
{
    private const int StreamWriterBufferSize = 256 * 1024;
    private const string ParameterSetInputObjectPathDelimiter = "InputObjectPathDelimiter";
    private const string ParameterSetInputObjectPathCulture = "InputObjectPathCulture";
    private const string ParameterSetInputObjectLiteralPathDelimiter = "InputObjectLiteralPathDelimiter";
    private const string ParameterSetInputObjectLiteralPathCulture = "InputObjectLiteralPathCulture";
    private const string ParameterSetDocumentPathDelimiter = "DocumentPathDelimiter";
    private const string ParameterSetDocumentPathCulture = "DocumentPathCulture";
    private const string ParameterSetDocumentLiteralPathDelimiter = "DocumentLiteralPathDelimiter";
    private const string ParameterSetDocumentLiteralPathCulture = "DocumentLiteralPathCulture";
    private CsvObjectWriter? _streamingWriter;
    private readonly CsvPowerShellObjectProjector _objectProjector = new();
    private string? _resolvedPath;
    private string[]? _appendHeader;
    private bool _appendToExistingFile;
    private bool _skipOutput;
    private bool _wroteOutput;

    /// <summary>Objects to export into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    public object? InputObject { get; set; }

    /// <summary>CSV document to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Destination CSV path.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Alias("FilePath", "OutputPath", "OutPath")]
    public string? Path { get; set; }

    /// <summary>Literal destination CSV path. Wildcards are not expanded.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    [Alias("PSPath", "LP")]
    public string? LiteralPath { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectLiteralPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentLiteralPathDelimiter)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectLiteralPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentLiteralPathCulture)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Omit the header row from the output.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Override the newline sequence.</summary>
    [Parameter]
    public string? NewLine { get; set; }

    /// <summary>Culture used for value formatting.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when writing files.</summary>
    [Parameter]
    public Encoding? Encoding { get; set; }

    /// <summary>Controls how formula-like values are written.</summary>
    [Parameter]
    public CsvFormulaInjectionPolicy FormulaInjectionPolicy { get; set; } = CsvFormulaInjectionPolicy.Preserve;

    /// <summary>Controls when CSV fields are quoted. Defaults to quoting only fields that need it.</summary>
    [Parameter]
    public CsvQuoteMode UseQuotes { get; set; } = CsvQuoteMode.AsNeeded;

    /// <summary>Field names that should always be quoted when <see cref="UseQuotes"/> is AsNeeded.</summary>
    [Parameter]
    public string[]? QuoteFields { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for the exported file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <summary>Append rows to an existing CSV file. Existing headers are reused when present.</summary>
    [Parameter]
    public SwitchParameter Append { get; set; }

    /// <summary>Do not overwrite an existing CSV file.</summary>
    [Parameter]
    [Alias("NoOverwrite")]
    public SwitchParameter NoClobber { get; set; }

    /// <summary>Allow overwriting read-only files and appending rows with missing existing columns.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        if (UseCulture.IsPresent)
        {
            var separator = (Culture ?? CultureInfo.CurrentCulture).TextInfo.ListSeparator;
            if (string.IsNullOrEmpty(separator) || separator.Length != 1)
            {
                throw new PSArgumentException("The selected culture must provide a single-character list separator.");
            }

            Delimiter = separator[0];
        }
    }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (IsDocumentParameterSet())
        {
            ExportDocument(Document);
            return;
        }

        WriteStreamingInputObject(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (_streamingWriter != null)
        {
            DisposeStreamingWriter();
            WritePassThru();
        }
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        DisposeStreamingWriter();
    }

    private void ExportDocument(CsvDocument document)
    {
        if (document == null || !TryPrepareOutput("Write CSV"))
        {
            return;
        }

        if (Append.IsPresent)
        {
            AppendDocument(document);
        }
        else
        {
            document.Save(_resolvedPath!, CreateSaveOptions());
        }

        _wroteOutput = true;
        WritePassThru();
    }

    private void AppendDocument(CsvDocument document)
    {
        var options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
        using var writer = CreateTextWriter(append: true, options);
        using var csvWriter = new CsvObjectWriter(writer, options);

        if (_appendHeader is { Length: > 0 })
        {
            ValidateDocumentAppendHeader(document, _appendHeader);
            WriteDocumentRows(document, csvWriter, _appendHeader, projectByName: true);
            return;
        }

        WriteDocumentRows(document, csvWriter, document.Header, projectByName: false);
    }

    private void WriteStreamingInputObject(object? value)
    {
        var writer = EnsureStreamingWriter();
        if (writer == null)
        {
            return;
        }

        try
        {
            _objectProjector.WriteObject(value, writer);
        }
        catch
        {
            DisposeStreamingWriter();
            throw;
        }
    }

    private CsvObjectWriter? EnsureStreamingWriter()
    {
        if (_streamingWriter != null)
        {
            return _streamingWriter;
        }

        if (!TryPrepareOutput("Write CSV"))
        {
            return null;
        }

        var options = CreateSaveOptions();
        if (Append.IsPresent)
        {
            options = CreateSaveOptions(includeHeader: !NoHeader.IsPresent && !_appendToExistingFile);
            if (_appendHeader is { Length: > 0 })
            {
                _objectProjector.UseColumns(_appendHeader, validateFirstRowColumns: !Force.IsPresent);
            }
        }

        var fileWriter = CreateTextWriter(Append.IsPresent, options);
        _streamingWriter = new CsvObjectWriter(fileWriter, options);
        _wroteOutput = true;
        return _streamingWriter;
    }

    private bool TryPrepareOutput(string action)
    {
        if (_skipOutput)
        {
            return false;
        }

        if (_wroteOutput)
        {
            WriteError(new ErrorRecord(
                new InvalidOperationException("Path can only be written once per invocation."),
                "CsvOutputAlreadyWritten",
                ErrorCategory.InvalidOperation,
                GetTargetPathForErrors()));
            _skipOutput = true;
            return false;
        }

        _resolvedPath = ResolveOutputPath();
        if (!ShouldProcess(_resolvedPath, action))
        {
            _skipOutput = true;
            return false;
        }

        var needsFileState = Append.IsPresent || NoClobber.IsPresent || Force.IsPresent;
        var fileExists = needsFileState && File.Exists(_resolvedPath);
        if (fileExists && NoClobber.IsPresent && !Append.IsPresent)
        {
            WriteError(new ErrorRecord(
                new IOException($"File '{_resolvedPath}' already exists."),
                "CsvFileExistsNoClobber",
                ErrorCategory.ResourceExists,
                _resolvedPath));
            _skipOutput = true;
            return false;
        }

        var directory = System.IO.Path.GetDirectoryName(_resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        if (fileExists && Force.IsPresent)
        {
            var fileInfo = new FileInfo(_resolvedPath);
            if (fileInfo.IsReadOnly)
            {
                fileInfo.IsReadOnly = false;
            }
        }

        _appendToExistingFile = Append.IsPresent && fileExists && new FileInfo(_resolvedPath).Length > 0;
        _appendHeader = _appendToExistingFile && !NoHeader.IsPresent
            ? ReadAppendHeader(_resolvedPath)
            : null;
        if (_appendToExistingFile)
        {
            EnsureAppendStartsOnNewRecord(_resolvedPath);
        }

        return true;
    }

    private void DisposeStreamingWriter()
    {
        _streamingWriter?.Dispose();
        _streamingWriter = null;
        _appendHeader = null;
        _appendToExistingFile = false;
        _objectProjector.Reset();
    }

    private void WritePassThru()
    {
        if (PassThru.IsPresent && !string.IsNullOrWhiteSpace(_resolvedPath))
        {
            WriteObject(new FileInfo(_resolvedPath!));
        }
    }

    private CsvSaveOptions CreateSaveOptions(bool? includeHeader = null)
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            IncludeHeader = includeHeader ?? !NoHeader.IsPresent,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            Encoding = Encoding,
            FormulaInjectionPolicy = FormulaInjectionPolicy,
            QuoteMode = UseQuotes,
            QuoteFields = QuoteFields
        };

        if (!string.IsNullOrEmpty(NewLine))
        {
            options.NewLine = NewLine!;
        }

        return options;
    }

    private bool IsDocumentParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetDocumentPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathCulture, StringComparison.OrdinalIgnoreCase);

    private bool IsLiteralPathParameterSet() =>
        string.Equals(ParameterSetName, ParameterSetInputObjectLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetInputObjectLiteralPathCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentLiteralPathCulture, StringComparison.OrdinalIgnoreCase);

    private string ResolveOutputPath()
    {
        var path = IsLiteralPathParameterSet()
            ? LiteralPath
            : Path;

        if (string.IsNullOrWhiteSpace(path))
        {
            throw new PSArgumentException("A destination path is required.");
        }

        return SessionState.Path.GetUnresolvedProviderPathFromPSPath(path);
    }

    private string? GetTargetPathForErrors() => IsLiteralPathParameterSet() ? LiteralPath : Path;

    private StreamWriter CreateTextWriter(bool append, CsvSaveOptions options)
    {
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var mode = append ? FileMode.Append : FileMode.Create;
        var stream = new FileStream(_resolvedPath!, mode, FileAccess.Write, FileShare.Read, StreamWriterBufferSize, FileOptions.SequentialScan);
        return new StreamWriter(stream, encoding, bufferSize: StreamWriterBufferSize);
    }

    private void EnsureAppendStartsOnNewRecord(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.Read, StreamWriterBufferSize, FileOptions.SequentialScan);
        if (stream.Length == 0)
        {
            return;
        }

        stream.Position = stream.Length - 1;
        var lastByte = stream.ReadByte();
        if (lastByte == '\n' || lastByte == '\r')
        {
            return;
        }

        stream.Position = stream.Length;
        var encoding = (Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        var newLineBytes = encoding.GetBytes(CreateSaveOptions().NewLine);
        stream.Write(newLineBytes, 0, newLineBytes.Length);
    }

    private string[] ReadAppendHeader(string path)
    {
        var options = new CsvLoadOptions
        {
            Delimiter = Delimiter,
            Encoding = Encoding,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            Mode = CsvLoadMode.Stream
        };

        return CsvDocument.Load(path, options).Header.ToArray();
    }

    private void ValidateDocumentAppendHeader(CsvDocument document, IReadOnlyList<string> appendHeader)
    {
        if (Force.IsPresent)
        {
            return;
        }

        var documentHeader = new HashSet<string>(document.Header, StringComparer.OrdinalIgnoreCase);
        foreach (var column in appendHeader)
        {
            if (!documentHeader.Contains(column))
            {
                throw new CsvException($"Cannot append CSV because the document is missing the existing column '{column}'. Use -Force to append with blank values for missing columns.");
            }
        }
    }

    private static void WriteDocumentRows(CsvDocument document, CsvObjectWriter writer, IReadOnlyList<string> columns, bool projectByName)
    {
        foreach (var row in document.AsEnumerable())
        {
            writer.WriteRow(
                columns,
                columns.Count,
                (Row: row, Columns: columns, ProjectByName: projectByName),
                static (state, index) => state.ProjectByName
                    ? TryGetRowValue(state.Row, state.Columns[index])
                    : index < state.Row.FieldCount ? state.Row[index] : null);
        }
    }

    private static object? TryGetRowValue(CsvRow row, string column)
    {
        try
        {
            return row[column];
        }
        catch
        {
            return null;
        }
    }
}
