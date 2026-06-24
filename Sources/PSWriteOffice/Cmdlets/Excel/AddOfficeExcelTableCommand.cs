using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;
using PSWriteOffice.Services.Table;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes tabular data to the current worksheet and formats it as an Excel table.</summary>
/// <para>Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.</para>
/// <example>
///   <summary>Insert a table starting at A1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
///   ExcelSheet 'Data' { Add-OfficeExcelTable -InputObject $data -TableName 'Sales' }</code>
///   <para>Writes two rows and formats them as a styled Excel table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTable")]
[Alias("ExcelTable")]
public sealed class AddOfficeExcelTableCommand : PSCmdlet
{
    private readonly List<object?> _items = new();

    /// <summary>Source objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences to convert into table rows.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("Data", "DataTable")]
    public object? InputObject { get; set; }

    /// <summary>Starting row for the data (1-based).</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Starting column for the data (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Skip writing headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Projection to apply before writing the table.</summary>
    [Parameter]
    public OfficeTableView View { get; set; } = OfficeTableView.Normal;

    /// <summary>Name to assign to the table.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public string TableStyle { get; set; } = "TableStyleMedium9";

    /// <summary>Emphasize the first table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowFirstColumn { get; set; }

    /// <summary>Emphasize the last table column when the selected style supports it.</summary>
    [Parameter]
    public SwitchParameter ShowLastColumn { get; set; }

    /// <summary>Disable alternating row stripes for the created table.</summary>
    [Parameter]
    public SwitchParameter NoRowStripes { get; set; }

    /// <summary>Enable alternating column stripes for the created table.</summary>
    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    /// <summary>Disable AutoFilter dropdowns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Auto-fit the table columns after insertion.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Return the created range string.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        TableInputCollector.AddInput(_items, InputObject, preserveTabularInput: true);
    }

    /// <inheritdoc />
    protected override void EndProcessing()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        var rows = TableInputCollector.RequireRows(_items, nameof(InputObject));
        var projectedRows = TableViewProjection.Project(rows, View);
        var table = ExcelTabularInputService.ToDataTable(projectedRows, TableName);
        if (table.Columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer columns from the supplied data.");
        }

        if (StartRow < 1 || StartColumn < 1)
        {
            throw new ArgumentOutOfRangeException("StartRow/StartColumn must be 1 or greater.");
        }

        if (!Enum.TryParse(TableStyle, ignoreCase: true, out TableStyle style))
        {
            throw new PSArgumentException($"Unknown table style '{TableStyle}'.", nameof(TableStyle));
        }

        var resolvedTableName = ResolveTableName(table);
        var range = sheet.InsertDataTableAsTable(
            table,
            startRow: StartRow,
            startColumn: StartColumn,
            includeHeaders: !NoHeader.IsPresent,
            tableName: resolvedTableName,
            style: style,
            includeAutoFilter: !NoAutoFilter.IsPresent);
        ExcelTableStyleOptionService.Apply(
            sheet,
            range,
            style,
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowFirstColumn), ShowFirstColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowLastColumn), ShowLastColumn),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(NoRowStripes), NoRowStripes),
            ExcelTableStyleOptionService.IsSwitchPresent(this, nameof(ShowColumnStripes), ShowColumnStripes));
        context.RegisterTableRange(sheet, resolvedTableName, range);

        if (AutoFit.IsPresent)
        {
            var columnIndexes = Enumerable.Range(StartColumn, table.Columns.Count);
            sheet.AutoFitColumnsFor(columnIndexes);
        }

        if (PassThru.IsPresent)
        {
            WriteObject(range);
        }
    }

    private string? ResolveTableName(System.Data.DataTable table)
    {
        if (!string.IsNullOrWhiteSpace(TableName))
        {
            return TableName;
        }

        return string.IsNullOrWhiteSpace(table.TableName) ? null : table.TableName;
    }
}
