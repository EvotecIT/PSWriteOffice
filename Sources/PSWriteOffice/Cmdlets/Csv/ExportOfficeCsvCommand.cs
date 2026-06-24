using System;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;
using PSWriteOffice.Services;

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
    private const int StreamWriterBufferSize = 32 * 1024;
    private const string ParameterSetInputObjectPathDelimiter = "InputObjectPathDelimiter";
    private const string ParameterSetInputObjectPathCulture = "InputObjectPathCulture";
    private const string ParameterSetInputObjectPathDelimiterQuoteFields = "InputObjectPathDelimiterQuoteFields";
    private const string ParameterSetInputObjectPathCultureQuoteFields = "InputObjectPathCultureQuoteFields";
    private const string ParameterSetDocumentPathDelimiter = "DocumentPathDelimiter";
    private const string ParameterSetDocumentPathCulture = "DocumentPathCulture";
    private const string ParameterSetDocumentPathDelimiterQuoteFields = "DocumentPathDelimiterQuoteFields";
    private const string ParameterSetDocumentPathCultureQuoteFields = "DocumentPathCultureQuoteFields";
    private CsvObjectWriter? _streamingWriter;
    private string[]? _streamingColumns;
    private object?[]? _streamingValues;
    private string? _resolvedPath;
    private bool _skipOutput;
    private bool _wroteOutput;

    /// <summary>Objects to export into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathDelimiterQuoteFields)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectPathCultureQuoteFields)]
    public object? InputObject { get; set; }

    /// <summary>CSV document to export.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentPathCultureQuoteFields)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Destination CSV path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "OutputPath", "OutPath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiterQuoteFields)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiterQuoteFields)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathCultureQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathCultureQuoteFields)]
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

    /// <summary>Controls when CSV fields are quoted.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectPathCulture)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentPathCulture)]
    public CsvQuoteMode UseQuotes { get; set; } = CsvQuoteMode.Always;

    /// <summary>Field names that should always be quoted when <see cref="UseQuotes"/> is AsNeeded.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectPathCultureQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentPathCultureQuoteFields)]
    public string[]? QuoteFields { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for the exported file.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

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

        document.Save(_resolvedPath!, CreateSaveOptions());
        _wroteOutput = true;
        WritePassThru();
    }

    private void WriteStreamingInputObject(object? value)
    {
        var writer = EnsureStreamingWriter();
        if (writer == null)
        {
            return;
        }

        if (_streamingColumns != null && _streamingValues != null &&
            PowerShellObjectNormalizer.TryProjectItemInto(value, _streamingColumns, _streamingValues))
        {
            writer.WriteRow(_streamingColumns, _streamingValues);
            return;
        }

        if (PowerShellObjectNormalizer.TryProjectItem(value, null, out var columns, out var values))
        {
            _streamingColumns = columns;
            _streamingValues = new object?[columns.Length];
            writer.WriteRow(_streamingColumns, values);
            return;
        }

        writer.WriteObject(PowerShellObjectNormalizer.NormalizeItem(value));
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
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var fileWriter = new StreamWriter(_resolvedPath!, append: false, encoding, bufferSize: StreamWriterBufferSize);
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
                Path));
            _skipOutput = true;
            return false;
        }

        _resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        if (!ShouldProcess(_resolvedPath, action))
        {
            _skipOutput = true;
            return false;
        }

        var directory = System.IO.Path.GetDirectoryName(_resolvedPath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return true;
    }

    private void DisposeStreamingWriter()
    {
        _streamingWriter?.Dispose();
        _streamingWriter = null;
        _streamingValues = null;
    }

    private void WritePassThru()
    {
        if (PassThru.IsPresent && !string.IsNullOrWhiteSpace(_resolvedPath))
        {
            WriteObject(new FileInfo(_resolvedPath!));
        }
    }

    private CsvSaveOptions CreateSaveOptions()
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            IncludeHeader = !NoHeader.IsPresent,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            Encoding = Encoding,
            FormulaInjectionPolicy = FormulaInjectionPolicy,
            QuoteMode = QuoteFields is { Length: > 0 } ? CsvQuoteMode.AsNeeded : UseQuotes,
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
        string.Equals(ParameterSetName, ParameterSetDocumentPathDelimiterQuoteFields, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentPathCultureQuoteFields, StringComparison.OrdinalIgnoreCase);
}
