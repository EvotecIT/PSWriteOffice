using System;
using System.Collections.Generic;
using System.Management.Automation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds a pivot table to a worksheet.</summary>
/// <example>
///   <summary>Create a basic pivot table.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Add-OfficeExcelPivotTable -SourceRange 'A1:D200' -DestinationCell 'F2' -RowField 'Region' -DataField 'Sales'</code>
///   <para>Creates a pivot table in F2 using Region as rows and Sales as the data field.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelPivotTable", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelPivotTable")]
public sealed class AddOfficeExcelPivotTableCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Source data range including header row (e.g., "A1:D200").</summary>
    [Parameter(Mandatory = true)]
    public string SourceRange { get; set; } = string.Empty;

    /// <summary>Top-left destination cell for the pivot table (e.g., "F2").</summary>
    [Parameter(Mandatory = true)]
    public string DestinationCell { get; set; } = string.Empty;

    /// <summary>Optional pivot table name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Row fields (header names).</summary>
    [Parameter]
    public string[]? RowField { get; set; }

    /// <summary>Column fields (header names).</summary>
    [Parameter]
    public string[]? ColumnField { get; set; }

    /// <summary>Page fields (header names) used as filters.</summary>
    [Parameter]
    public string[]? PageField { get; set; }

    /// <summary>Data fields (header names). Defaults to the last column when omitted.</summary>
    [Parameter]
    public string[]? DataField { get; set; }

    /// <summary>Aggregation functions (Sum, Count, Average, etc.).</summary>
    [Parameter]
    public string[]? DataFunction { get; set; }

    /// <summary>Disable row grand totals.</summary>
    [Parameter]
    public SwitchParameter NoRowGrandTotals { get; set; }

    /// <summary>Disable column grand totals.</summary>
    [Parameter]
    public SwitchParameter NoColumnGrandTotals { get; set; }

    /// <summary>Optional pivot table style name.</summary>
    [Parameter]
    public string? PivotStyle { get; set; }

    /// <summary>Pivot layout (Compact, Outline, Tabular).</summary>
    [Parameter]
    public string Layout { get; set; } = "Compact";

    /// <summary>Show data fields on rows.</summary>
    [Parameter]
    public SwitchParameter DataOnRows { get; set; }

    /// <summary>Show data fields on columns.</summary>
    [Parameter]
    public SwitchParameter DataOnColumns { get; set; }

    /// <summary>Show field headers.</summary>
    [Parameter]
    public SwitchParameter ShowHeaders { get; set; }

    /// <summary>Hide field headers.</summary>
    [Parameter]
    public SwitchParameter HideHeaders { get; set; }

    /// <summary>Show empty rows.</summary>
    [Parameter]
    public SwitchParameter ShowEmptyRows { get; set; }

    /// <summary>Hide empty rows.</summary>
    [Parameter]
    public SwitchParameter HideEmptyRows { get; set; }

    /// <summary>Show empty columns.</summary>
    [Parameter]
    public SwitchParameter ShowEmptyColumns { get; set; }

    /// <summary>Hide empty columns.</summary>
    [Parameter]
    public SwitchParameter HideEmptyColumns { get; set; }

    /// <summary>Show drill indicators.</summary>
    [Parameter]
    public SwitchParameter ShowDrill { get; set; }

    /// <summary>Hide drill indicators.</summary>
    [Parameter]
    public SwitchParameter HideDrill { get; set; }

    /// <summary>Emit the worksheet after creating the pivot table.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var sheet = ResolveSheet();
        var dataFields = BuildDataFields();
        if (!Enum.TryParse(Layout, ignoreCase: true, out ExcelPivotLayout layout))
        {
            throw new PSArgumentException($"Unknown pivot layout '{Layout}'.", nameof(Layout));
        }

        var dataOnRows = ResolveToggle(DataOnRows, DataOnColumns, "DataOnRows/DataOnColumns");
        var showHeaders = ResolveToggle(ShowHeaders, HideHeaders, "ShowHeaders/HideHeaders");
        var showEmptyRows = ResolveToggle(ShowEmptyRows, HideEmptyRows, "ShowEmptyRows/HideEmptyRows");
        var showEmptyColumns = ResolveToggle(ShowEmptyColumns, HideEmptyColumns, "ShowEmptyColumns/HideEmptyColumns");
        var showDrill = ResolveToggle(ShowDrill, HideDrill, "ShowDrill/HideDrill");

        sheet.AddPivotTable(
            sourceRange: SourceRange,
            destinationCell: DestinationCell,
            name: Name,
            rowFields: RowField,
            columnFields: ColumnField,
            pageFields: PageField,
            dataFields: dataFields,
            showRowGrandTotals: !NoRowGrandTotals.IsPresent,
            showColumnGrandTotals: !NoColumnGrandTotals.IsPresent,
            pivotStyleName: PivotStyle,
            layout: layout,
            dataOnRows: dataOnRows,
            showHeaders: showHeaders,
            showEmptyRows: showEmptyRows,
            showEmptyColumns: showEmptyColumns,
            showDrill: showDrill);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private IEnumerable<ExcelPivotDataField>? BuildDataFields()
    {
        if (DataField == null || DataField.Length == 0)
        {
            if (DataFunction != null && DataFunction.Length > 0)
            {
                throw new PSArgumentException("DataFunction requires DataField to be provided.");
            }
            return null;
        }

        var functions = ParseFunctions(DataFunction);
        if (functions.Count > 1 && functions.Count != DataField.Length)
        {
            throw new PSArgumentException("When providing multiple DataFunction values, the count must match DataField.");
        }

        var result = new List<ExcelPivotDataField>(DataField.Length);
        for (int i = 0; i < DataField.Length; i++)
        {
            var fieldName = DataField[i];
            if (string.IsNullOrWhiteSpace(fieldName)) continue;

            var fn = functions.Count switch
            {
                0 => DataConsolidateFunctionValues.Sum,
                1 => functions[0],
                _ => functions[i]
            };
            result.Add(new ExcelPivotDataField(fieldName.Trim(), fn));
        }

        return result.Count == 0 ? null : result;
    }

    private static List<DataConsolidateFunctionValues> ParseFunctions(string[]? functions)
    {
        var result = new List<DataConsolidateFunctionValues>();
        if (functions == null || functions.Length == 0) return result;

        foreach (var raw in functions)
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            if (!Enum.TryParse(raw, ignoreCase: true, out DataConsolidateFunctionValues fn))
            {
                throw new PSArgumentException($"Unknown DataFunction '{raw}'.");
            }
            result.Add(fn);
        }

        return result;
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }

    private static bool? ResolveToggle(SwitchParameter enable, SwitchParameter disable, string name)
    {
        if (enable.IsPresent && disable.IsPresent)
        {
            throw new PSArgumentException($"Cannot set both {name}.");
        }

        if (enable.IsPresent) return true;
        if (disable.IsPresent) return false;
        return null;
    }
}
