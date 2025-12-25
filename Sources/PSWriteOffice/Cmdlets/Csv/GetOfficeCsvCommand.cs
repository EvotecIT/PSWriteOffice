using System.Globalization;
using System.IO;
using System.Management.Automation;
using System.Text;
using OfficeIMO.CSV;

namespace PSWriteOffice.Cmdlets.Csv;

/// <summary>Loads a CSV document from disk or parses CSV text.</summary>
/// <para>Returns an <see cref="CsvDocument"/> for inspection or further transformations.</para>
/// <example>
///   <summary>Load a CSV file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$csv = Get-OfficeCsv -Path .\data.csv</code>
///   <para>Loads the CSV file into an OfficeIMO CsvDocument.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficeCsv", DefaultParameterSetName = ParameterSetPath)]
[OutputType(typeof(CsvDocument))]
public sealed class GetOfficeCsvCommand : PSCmdlet
{
    private const string ParameterSetPath = "Path";
    private const string ParameterSetText = "Text";

    /// <summary>Path to the CSV file.</summary>
    [Parameter(Mandatory = true, Position = 0, ParameterSetName = ParameterSetPath)]
    [Alias("FilePath", "Path")]
    public string InputPath { get; set; } = string.Empty;

    /// <summary>CSV text to parse.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetText)]
    public string Text { get; set; } = string.Empty;

    /// <summary>Indicates whether the first record is a header row.</summary>
    [Parameter]
    public bool HasHeaderRow { get; set; } = true;

    /// <summary>Field delimiter character.</summary>
    [Parameter]
    public char Delimiter { get; set; } = ',';

    /// <summary>Trim whitespace around unquoted fields.</summary>
    [Parameter]
    public bool TrimWhitespace { get; set; } = true;

    /// <summary>Allow empty lines in the input.</summary>
    [Parameter]
    public SwitchParameter AllowEmptyLines { get; set; }

    /// <summary>Load mode controlling materialization.</summary>
    [Parameter]
    public CsvLoadMode Mode { get; set; } = CsvLoadMode.InMemory;

    /// <summary>Culture used for type conversions.</summary>
    [Parameter]
    public CultureInfo? Culture { get; set; }

    /// <summary>Encoding used when reading the file.</summary>
    [Parameter]
    public Encoding? Encoding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var options = new CsvLoadOptions
        {
            HasHeaderRow = HasHeaderRow,
            Delimiter = Delimiter,
            TrimWhitespace = TrimWhitespace,
            AllowEmptyLines = AllowEmptyLines.IsPresent,
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

        CsvDocument document;
        if (ParameterSetName == ParameterSetPath)
        {
            var resolvedPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException($"File '{resolvedPath}' was not found.", resolvedPath);
            }

            document = CsvDocument.Load(resolvedPath, options);
        }
        else
        {
            document = CsvDocument.Parse(Text ?? string.Empty, options);
        }

        WriteObject(document);
    }
}
