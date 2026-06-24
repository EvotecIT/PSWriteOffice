using System;
using System.Collections.Generic;
using System.Globalization;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Converts objects or a CSV document into CSV text.</summary>
/// <para>Use <c>Export-OfficeCsv</c> when the destination is a file.</para>
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
///   $csv = $rows | ConvertTo-OfficeCsv -Delimiter ';'</code>
///   <para>Uses ordered dictionaries to enforce column order and a custom delimiter.</para>
/// </example>
/// <example>
///   <summary>Write CSV without headers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = $data | ConvertTo-OfficeCsv -NoHeader</code>
///   <para>Writes rows only when a downstream system expects headerless CSV.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeCsv", DefaultParameterSetName = ParameterSetInputObjectDelimiter)]
[OutputType(typeof(string))]
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

    /// <summary>Omit the header row from the output.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Override the newline sequence.</summary>
    [Parameter]
    public string? NewLine { get; set; }

    /// <summary>Culture used for value formatting.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding carried into the CSV save options.</summary>
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

    /// <inheritdoc />
    protected override void BeginProcessing()
    {
        ApplyCultureDelimiter();
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

        _items.Add(InputObject);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        if (IsDocumentParameterSet() || _items.Count == 0)
        {
            return;
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(_items);
        var document = CsvDocument.FromObjects(normalized, Delimiter, Culture, Encoding);
        EmitCsv(document);
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

    private void EmitCsv(CsvDocument document)
    {
        WriteObject(document.ToString(CreateSaveOptions()));
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
        string.Equals(ParameterSetName, ParameterSetDocumentDelimiter, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentCulture, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentDelimiterQuoteFields, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(ParameterSetName, ParameterSetDocumentCultureQuoteFields, StringComparison.OrdinalIgnoreCase);
}
