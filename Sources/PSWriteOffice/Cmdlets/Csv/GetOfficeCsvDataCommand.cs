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
[Cmdlet(VerbsCommon.Get, "OfficeCsvData")]
public sealed class GetOfficeCsvDataCommand : PSCmdlet
{
    /// <summary>CSV document to read when already loaded.</summary>
    [Parameter(ValueFromPipeline = true)]
    public CsvDocument? Document { get; set; }

    /// <summary>Path to a CSV file.</summary>
    [Parameter(Position = 0)]
    [Alias("FilePath", "Path")]
    public string? InputPath { get; set; }

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

    /// <summary>Emit dictionaries instead of PSCustomObjects.</summary>
    [Parameter]
    public SwitchParameter AsHashtable { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var document = Document;

        if (document == null)
        {
            if (string.IsNullOrWhiteSpace(InputPath))
            {
                throw new PSArgumentException("Specify -Path or provide a CsvDocument on the pipeline.");
            }

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

            var resolved = SessionState.Path.GetUnresolvedProviderPathFromPSPath(InputPath);
            if (!File.Exists(resolved))
            {
                throw new FileNotFoundException($"File '{resolved}' was not found.", resolved);
            }

            document = CsvDocument.Load(resolved, options);
        }

        var header = document.Header;
        foreach (var row in document.AsEnumerable())
        {
            var rowValues = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < header.Count && i < row.FieldCount; i++)
            {
                rowValues[header[i]] = row[i];
            }

            if (AsHashtable.IsPresent)
            {
                WriteObject(rowValues);
            }
            else
            {
                var psObj = new PSObject();
                foreach (var kv in rowValues)
                {
                    psObj.Properties.Add(new PSNoteProperty(kv.Key, kv.Value));
                }
                WriteObject(psObj);
            }
        }
    }
}
