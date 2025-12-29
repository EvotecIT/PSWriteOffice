using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services;

namespace PSWriteOffice.Cmdlets.PowerPoint;

/// <summary>Adds a table to a PowerPoint slide.</summary>
/// <para>Builds a table from data rows or creates a blank grid with a fixed size.</para>
/// <example>
///   <summary>Create a table from objects.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$rows = @([pscustomobject]@{ Item='Alpha'; Qty=2 }, [pscustomobject]@{ Item='Beta'; Qty=4 })
///   Add-OfficePowerPointTable -Slide $slide -Data $rows -X 60 -Y 140 -Width 420 -Height 200</code>
///   <para>Creates a table with headers and two data rows.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficePowerPointTable", DefaultParameterSetName = ParameterSetData)]
public sealed class AddOfficePowerPointTableCommand : PSCmdlet
{
    private const string ParameterSetData = "Data";
    private const string ParameterSetSize = "Size";

    /// <summary>Target slide that will receive the table.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, Position = 0)]
    public PowerPointSlide Slide { get; set; } = null!;

    /// <summary>Source objects to convert into table rows.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetData)]
    public object[] Data { get; set; } = Array.Empty<object>();

    /// <summary>Optional header order to apply to the table.</summary>
    [Parameter(ParameterSetName = ParameterSetData)]
    public string[]? Headers { get; set; }

    /// <summary>Skip writing header row.</summary>
    [Parameter(ParameterSetName = ParameterSetData)]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Row count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Rows { get; set; }

    /// <summary>Column count for an empty table.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetSize)]
    public int Columns { get; set; }

    /// <summary>Left offset (in points) from the slide origin.</summary>
    [Parameter]
    public double X { get; set; } = 0;

    /// <summary>Top offset (in points) from the slide origin.</summary>
    [Parameter]
    public double Y { get; set; } = 0;

    /// <summary>Table width in points.</summary>
    [Parameter]
    public double Width { get; set; } = 400;

    /// <summary>Table height in points.</summary>
    [Parameter]
    public double Height { get; set; } = 240;

    /// <summary>Optional table style ID (GUID string).</summary>
    [Parameter]
    public string? StyleId { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            ValidateDimensions();

            PowerPointTable table = ParameterSetName == ParameterSetSize
                ? CreateSizedTable()
                : CreateDataTable();

            if (!string.IsNullOrWhiteSpace(StyleId))
            {
                table.StyleId = StyleId;
            }

            WriteObject(table);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "PowerPointAddTableFailed", ErrorCategory.InvalidOperation, Slide));
        }
    }

    private PowerPointTable CreateSizedTable()
    {
        if (Rows <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(Rows), "Rows must be greater than 0.");
        }

        if (Columns <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(Columns), "Columns must be greater than 0.");
        }

        return Slide.AddTablePoints(Rows, Columns, X, Y, Width, Height);
    }

    private PowerPointTable CreateDataTable()
    {
        if (Data == null || Data.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data row.", nameof(Data));
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItems(Data);
        var rows = NormalizeRows(normalized);
        var headers = ResolveHeaders(rows);

        if (headers.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer columns from the supplied data.");
        }

        var columns = headers
            .Select(header => PowerPointTableColumn<Dictionary<string, object?>>.Create(
                header,
                row => row.TryGetValue(header, out var value) ? value : null))
            .ToList();

        return Slide.AddTablePoints(rows, columns, includeHeaders: !NoHeader.IsPresent, X, Y, Width, Height);
    }

    private void ValidateDimensions()
    {
        if (Width <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(Width), "Width must be greater than 0.");
        }

        if (Height <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(Height), "Height must be greater than 0.");
        }
    }

    private List<Dictionary<string, object?>> NormalizeRows(IReadOnlyList<object?> items)
    {
        var rows = new List<Dictionary<string, object?>>(items.Count);
        foreach (var item in items)
        {
            if (item == null)
            {
                rows.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase));
                continue;
            }

            if (item is IDictionary dict)
            {
                var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in dict)
                {
                    var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        continue;
                    }
                    row[key] = entry.Value;
                }

                rows.Add(row);
                continue;
            }

            rows.Add(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Value"] = item
            });
        }

        return rows;
    }

    private List<string> ResolveHeaders(IReadOnlyList<Dictionary<string, object?>> rows)
    {
        if (Headers != null && Headers.Length > 0)
        {
            var explicitHeaders = Headers
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (explicitHeaders.Count == 0)
            {
                throw new PSArgumentException("Headers cannot be empty.", nameof(Headers));
            }

            return explicitHeaders;
        }

        var headers = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var row in rows)
        {
            foreach (var key in row.Keys)
            {
                if (seen.Add(key))
                {
                    headers.Add(key);
                }
            }
        }

        return headers;
    }
}
