using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;
using PSWriteOffice.Services.Csv;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts objects or a CSV document into CSV text or a file.</summary>
/// <para>By default returns CSV text; use <c>-OutputPath</c> to save to disk.</para>
/// <example>
///   <summary>Convert objects to CSV text.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = $data | ConvertTo-OfficeCsv</code>
///   <para>Generates CSV text from the input objects.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObject)]
[OutputType(typeof(string), typeof(FileInfo))]
public sealed class ConvertToOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetInputObject = "InputObject";
    private const string ParameterSetDocument = "Document";
    private readonly List<object?> _items = new();
    private bool _wroteOutputPath;

    /// <summary>CSV document to serialize.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public CsvDocument Document { get; set; } = null!;

    /// <summary>Objects to convert into CSV rows.</summary>
    [Parameter(ValueFromPipeline = true, ParameterSetName = ParameterSetInputObject)]
    public object? InputObject { get; set; }

    /// <summary>Field delimiter character.</summary>
    [Parameter]
    public char Delimiter { get; set; } = ',';

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

    /// <summary>Optional output path for the CSV file.</summary>
    [Parameter]
    [Alias("OutPath")]
    public string? OutputPath { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> when saving to disk.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document != null)
            {
                EmitCsv(Document);
            }
            return;
        }

        _items.Add(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            return;
        }

        if (_items.Count == 0)
        {
            return;
        }

        var document = CsvDocumentBuilder.FromObjects(_items, Delimiter, Culture, Encoding);
        EmitCsv(document);
    }

    private void EmitCsv(CsvDocument document)
    {
        var options = new CsvSaveOptions
        {
            Delimiter = Delimiter,
            IncludeHeader = IncludeHeader,
            Culture = Culture ?? CultureInfo.InvariantCulture,
            Encoding = Encoding
        };

        if (!string.IsNullOrEmpty(NewLine))
        {
            options.NewLine = NewLine;
        }

        if (!string.IsNullOrWhiteSpace(OutputPath))
        {
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
            var directory = Path.GetDirectoryName(resolved);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            document.Save(resolved, options);
            _wroteOutputPath = true;

            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(resolved));
            }
        }
        else
        {
            WriteObject(document.ToString(options));
        }
    }
}
