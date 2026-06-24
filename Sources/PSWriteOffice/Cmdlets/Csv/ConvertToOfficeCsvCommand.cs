using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts objects or a CSV document into CSV text or a file.</summary>
/// <para>By default returns CSV text; use <c>-OutputPath</c> to save to disk.</para>
/// <example>
///   <summary>Convert objects to CSV text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = $data | ConvertTo-OfficeCsv</code>
///   <para>Generates CSV text from the input objects.</para>
/// </example>
/// <example>
///   <summary>Export with a stable schema order.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @(
///     [ordered]@{ Id = 1; Name = 'Alpha'; Total = 10.5 },
///     [ordered]@{ Id = 2; Name = 'Beta'; Total = 7.25 }
///   )
///   $rows | ConvertTo-OfficeCsv -OutputPath .\export.csv -Delimiter ';'</code>
///   <para>Uses ordered dictionaries to enforce column order and a custom delimiter.</para>
/// </example>
/// <example>
///   <summary>Write CSV without headers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data | ConvertTo-OfficeCsv -IncludeHeader:$false -OutputPath .\noheader.csv</code>
///   <para>Writes rows only when a downstream system expects headerless CSV.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObjectDelimiter, SupportsShouldProcess = true)]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetInputObjectDelimiter = "InputObjectDelimiter";
    private const string ParameterSetInputObjectCulture = "InputObjectCulture";
    private const string ParameterSetDocumentDelimiter = "DocumentDelimiter";
    private const string ParameterSetDocumentCulture = "DocumentCulture";
    private const string ParameterSetInputObjectDelimiterQuoteFields = "InputObjectDelimiterQuoteFields";
    private const string ParameterSetInputObjectCultureQuoteFields = "InputObjectCultureQuoteFields";
    private const string ParameterSetDocumentDelimiterQuoteFields = "DocumentDelimiterQuoteFields";
    private const string ParameterSetDocumentCultureQuoteFields = "DocumentCultureQuoteFields";
    private readonly List<object?> _items = new();
    private CsvObjectWriter? _streamingWriter;
    private string[]? _streamingColumns;
    private object?[]? _streamingValues;
    private string? _streamingOutputPath;
    private bool _skipStreamingOutput;
    private bool _wroteOutputPath;

    /// <summary>CSV document to serialize.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentDelimiter)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentCulture)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocumentCultureQuoteFields)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Objects to convert into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectCulture)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectDelimiterQuoteFields)]
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObjectCultureQuoteFields)]
    public object? InputObject { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiterQuoteFields)]
    [Parameter(ParameterSetName = ParameterSetDocumentDelimiterQuoteFields)]
    public char Delimiter { get; set; } = ',';

    /// <summary>Use the list separator from the selected or current culture as the delimiter.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentCulture)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectCultureQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentCultureQuoteFields)]
    public SwitchParameter UseCulture { get; set; }

    /// <summary>Include the header row in the output.</summary>
    [Parameter]
    public bool IncludeHeader { get; set; } = true;

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
    [Parameter(ParameterSetName = ParameterSetInputObjectDelimiter)]
    [Parameter(ParameterSetName = ParameterSetInputObjectCulture)]
    [Parameter(ParameterSetName = ParameterSetDocumentDelimiter)]
    [Parameter(ParameterSetName = ParameterSetDocumentCulture)]
    public CsvQuoteMode UseQuotes { get; set; } = CsvQuoteMode.Always;

    /// <summary>Field names that should always be quoted when <see cref="UseQuotes"/> is AsNeeded.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetInputObjectCultureQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentDelimiterQuoteFields)]
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDocumentCultureQuoteFields)]
    public string[]? QuoteFields { get; set; }

    /// <summary>Optional output path for the CSV file.</summary>
    [Parameter]
    [Alias("Path", "OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
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
            if (Document != null)
            {
                EmitCsv(Document);
            }
            return;
        }

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteStreamingInputObject(InputObject);
            return;
        }

        _items.Add(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (IsDocumentParameterSet())
        {
            return;
        }

        if (_streamingWriter != null)
        {
            DisposeStreamingWriter();
            WriteStreamingPassThru();
            return;
        }

        if (_skipStreamingOutput)
        {
            return;
        }

        if (_items.Count == 0)
        {
            return;
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(_items);
        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            EmitObjectsCsv(normalized);
            return;
        }

        var document = CsvDocument.FromObjects(normalized, Delimiter, Culture, Encoding);
        EmitCsv(document);
    }

    /// <inheritdoc />
    protected override void StopProcessing()
    {
        DisposeStreamingWriter();
    }

    private void EmitCsv(CsvDocument document)
    {
        var options = CreateSaveOptions();

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
            WriteCsvFile(resolved => document.Save(resolved, options));
        }
        else
        {
            WriteObject(document.ToString(options));
        }
    }

    private void EmitObjectsCsv(IReadOnlyList<object?> items)
    {
        var options = CreateSaveOptions();
        WriteCsvFile(resolved => CsvDocument.SaveObjects(resolved, items, options));
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

        if (_skipStreamingOutput)
        {
            return null;
        }

        if (_wroteOutputPath)
        {
            WriteError(new ErrorRecord(
                new InvalidOperationException("OutputPath can only be used once per invocation."),
                "CsvOutputAlreadyWritten",
                ErrorCategory.InvalidOperation,
                OutputPath));
            _skipStreamingOutput = true;
            return null;
        }

        var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath!);
        if (!ShouldProcess(resolved, "Write CSV"))
        {
            _skipStreamingOutput = true;
            return null;
        }

        var directory = Path.GetDirectoryName(resolved);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var options = CreateSaveOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var fileWriter = new StreamWriter(resolved, append: false, encoding, bufferSize: 256 * 1024);
        _streamingWriter = new CsvObjectWriter(fileWriter, options);
        _streamingOutputPath = resolved;
        _wroteOutputPath = true;
        return _streamingWriter;
    }

    private void DisposeStreamingWriter()
    {
        _streamingWriter?.Dispose();
        _streamingWriter = null;
        _streamingValues = null;
    }

    private void WriteStreamingPassThru()
    {
        if (PassThru.IsPresent && !string.IsNullOrWhiteSpace(_streamingOutputPath))
        {
            WriteObject(new FileInfo(_streamingOutputPath!));
        }
    }

    private CsvSaveOptions CreateSaveOptions()
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            IncludeHeader = IncludeHeader,
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
        string.Equals(ParameterSetName, ParameterSetDocumentDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentDelimiterQuoteFields, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentCultureQuoteFields, StringComparison.OrdinalIgnoreCase);

    private void WriteCsvFile(Action<string> writeFile)
    {
        if (string.IsNullOrWhiteSpace(OutputPath))
        {
            return;
        }

        if (_wroteOutputPath)
        {
            WriteError(new ErrorRecord(
                new InvalidOperationException("OutputPath can only be used once per invocation."),
                "CsvOutputAlreadyWritten",
                ErrorCategory.InvalidOperation,
                OutputPath));
            return;
        }

        var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(OutputPath);
        if (!ShouldProcess(resolved, "Write CSV"))
        {
            return;
        }

        var directory = Path.GetDirectoryName(resolved);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        writeFile(resolved);
        _wroteOutputPath = true;

        if (PassThru.IsPresent)
        {
            WriteObject(new FileInfo(resolved));
        }
    }
}
