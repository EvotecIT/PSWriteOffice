using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a title block to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTitle")]
[Alias("ExcelReportTitle")]
public sealed class AddOfficeExcelReportTitleCommand : PSCmdlet
{
    /// <summary>Title text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Optional subtitle text.</summary>
    [Parameter(Position = 1)]
    public string? Subtitle { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().Title(Title, Subtitle);
    }
}

/// <summary>Adds a section heading to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSection")]
[Alias("ExcelReportSection")]
public sealed class AddOfficeExcelReportSectionCommand : PSCmdlet
{
    /// <summary>Section heading text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().Section(Text);
    }
}

/// <summary>Adds a paragraph line to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportParagraph")]
[Alias("ExcelReportParagraph")]
public sealed class AddOfficeExcelReportParagraphCommand : PSCmdlet
{
    /// <summary>Paragraph text.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Text { get; set; } = string.Empty;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().Paragraph(Text);
    }
}

/// <summary>Adds vertical spacing to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportSpacer")]
[Alias("ExcelReportSpacer")]
public sealed class AddOfficeExcelReportSpacerCommand : PSCmdlet
{
    /// <summary>Rows to advance. Defaults to the composer theme spacing.</summary>
    [Parameter(Position = 0)]
    public int Rows { get; set; } = -1;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().Spacer(Rows);
    }
}

/// <summary>Adds a colored callout block to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportCallout")]
[Alias("ExcelReportCallout")]
public sealed class AddOfficeExcelReportCalloutCommand : PSCmdlet
{
    /// <summary>Callout kind. Supported values include info, success, warning, error, and critical.</summary>
    [Parameter(Position = 0)]
    [ValidateSet("Info", "Success", "Warning", "Error", "Critical")]
    public string Kind { get; set; } = "Info";

    /// <summary>Callout title.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string Title { get; set; } = string.Empty;

    /// <summary>Callout body text.</summary>
    [Parameter(Mandatory = true, Position = 2)]
    public string Body { get; set; } = string.Empty;

    /// <summary>Width of the highlighted callout band in worksheet columns.</summary>
    [Parameter]
    public int WidthColumns { get; set; } = 8;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().Callout(Kind, Title, Body, WidthColumns);
    }
}

/// <summary>Adds a KPI row to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportKpiRow")]
[Alias("ExcelReportKpiRow")]
public sealed class AddOfficeExcelReportKpiRowCommand : PSCmdlet
{
    /// <summary>Hashtable or objects with Label/Value, Key/Value, Name/Value, or Title/Value properties.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public object Data { get; set; } = null!;

    /// <summary>Number of KPI cards per rendered row.</summary>
    [Parameter]
    public int PerRow { get; set; } = 3;

    /// <summary>Optional fill color for KPI labels.</summary>
    [Parameter]
    public string? LabelFillColor { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().KpiRow(ReportBlockInput.ToPairs(Data), PerRow, LabelFillColor);
    }
}

/// <summary>Adds a legend table to the current Excel report sheet.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportLegend")]
[Alias("ExcelReportLegend")]
public sealed class AddOfficeExcelReportLegendCommand : PSCmdlet
{
    /// <summary>Optional legend title.</summary>
    [Parameter(Position = 0)]
    public string? Title { get; set; }

    /// <summary>Column headers.</summary>
    [Parameter(Mandatory = true)]
    public string[] Headers { get; set; } = Array.Empty<string>();

    /// <summary>Rows. Each row may be an array, enumerable, hashtable, or object.</summary>
    [Parameter(Mandatory = true)]
    public object[] Rows { get; set; } = Array.Empty<object>();

    /// <summary>Optional first-column fill colors keyed by first-column value.</summary>
    [Parameter]
    public Hashtable? FirstColumnFillByValue { get; set; }

    /// <summary>Optional header fill color.</summary>
    [Parameter]
    public string? HeaderFillColor { get; set; }

    /// <summary>Use case-sensitive matching for first-column fill values.</summary>
    [Parameter]
    public SwitchParameter CaseSensitive { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        ExcelDslContext.Require(this).RequireComposer().SectionLegend(
            Title,
            Headers,
            Rows.Select(row => ReportBlockInput.ToRow(row, Headers)),
            ReportBlockInput.ToStringMap(FirstColumnFillByValue, CaseSensitive.IsPresent),
            HeaderFillColor);
    }
}

/// <summary>Adds an object table to the current Excel report sheet using the OfficeIMO sheet composer.</summary>
[Cmdlet(VerbsCommon.Add, "OfficeExcelReportTable")]
[Alias("ExcelReportTable")]
public sealed class AddOfficeExcelReportTableCommand : PSCmdlet
{
    /// <summary>Objects to flatten and render as a table.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public object[] Data { get; set; } = Array.Empty<object>();

    /// <summary>Optional section title displayed above the table.</summary>
    [Parameter(Position = 1)]
    public string? Title { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public string TableStyle { get; set; } = "TableStyleMedium9";

    /// <summary>Disable AutoFilter dropdowns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Do not freeze through the table header row.</summary>
    [Parameter]
    public SwitchParameter NoFreezeHeaderRow { get; set; }

    /// <summary>Disable composer auto-formatting for dynamic collection columns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFormatDynamicCollections { get; set; }

    /// <summary>Emit the A1 range used by the generated table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Data.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data row.", nameof(Data));
        }

        if (!Enum.TryParse(TableStyle, ignoreCase: true, out TableStyle style))
        {
            throw new PSArgumentException($"Unknown table style '{TableStyle}'.", nameof(TableStyle));
        }

        var composer = ExcelDslContext.Require(this).RequireComposer();
        var range = composer.TableFrom(
            Data,
            Title,
            style: style,
            autoFilter: !NoAutoFilter.IsPresent,
            freezeHeaderRow: !NoFreezeHeaderRow.IsPresent,
            visuals: options => options.AutoFormatDynamicCollections = !NoAutoFormatDynamicCollections.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(range);
        }
    }
}

internal static class ReportBlockInput
{
    public static IReadOnlyList<(string Label, object? Value)> ToPairs(object input)
    {
        if (input is IDictionary dictionary)
        {
            var pairs = new List<(string Label, object? Value)>();
            foreach (DictionaryEntry entry in dictionary)
            {
                pairs.Add((Convert.ToString(entry.Key, CultureInfo.InvariantCulture) ?? string.Empty, entry.Value));
            }

            return pairs;
        }

        var rows = input is IEnumerable enumerable and not string
            ? enumerable.Cast<object?>().Where(item => item != null).ToArray()
            : new[] { input };

        return rows.Select(item =>
        {
            var psObject = PSObject.AsPSObject(item);
            var label = GetProperty(psObject, "Label")
                ?? GetProperty(psObject, "Key")
                ?? GetProperty(psObject, "Name")
                ?? GetProperty(psObject, "Title")
                ?? Convert.ToString(item, CultureInfo.InvariantCulture)
                ?? string.Empty;
            var value = GetPropertyValue(psObject, "Value") ?? GetPropertyValue(psObject, "Count") ?? GetPropertyValue(psObject, "Total");
            return (label, value);
        }).ToArray();
    }

    public static IReadOnlyList<string> ToRow(object input, IReadOnlyList<string> headers)
    {
        if (input is IDictionary dictionary)
        {
            return headers.Select(header =>
            {
                foreach (DictionaryEntry entry in dictionary)
                {
                    if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), header, StringComparison.OrdinalIgnoreCase))
                    {
                        return Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? string.Empty;
                    }
                }

                return string.Empty;
            }).ToArray();
        }

        if (input is IEnumerable enumerable and not string)
        {
            return enumerable.Cast<object?>()
                .Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
                .ToArray();
        }

        var psObject = PSObject.AsPSObject(input);
        return headers.Select(header => Convert.ToString(GetPropertyValue(psObject, header), CultureInfo.InvariantCulture) ?? string.Empty).ToArray();
    }

    public static Dictionary<string, string>? ToStringMap(Hashtable? table, bool caseSensitive = false)
    {
        if (table == null || table.Count == 0)
        {
            return null;
        }

        var comparer = caseSensitive ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        var result = new Dictionary<string, string>(comparer);
        foreach (DictionaryEntry entry in table)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            var value = Convert.ToString(entry.Value, CultureInfo.InvariantCulture);
            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
            {
                result[key] = value;
            }
        }

        return result;
    }

    private static string? GetProperty(PSObject psObject, string name)
    {
        return Convert.ToString(GetPropertyValue(psObject, name), CultureInfo.InvariantCulture);
    }

    private static object? GetPropertyValue(PSObject psObject, string name)
    {
        return psObject.Properties[name]?.Value;
    }
}
