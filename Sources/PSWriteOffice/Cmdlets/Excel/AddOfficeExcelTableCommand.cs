using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes tabular data to the current worksheet and formats it as an Excel table.</summary>
/// <para>Accepts objects, dictionaries, DataTable/DataView/IDataReader inputs, or DataRow sequences and writes them into an Excel table with optional styling.</para>
/// <example>
///   <summary>Insert a table starting at A1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
///   ExcelSheet 'Data' { Add-OfficeExcelTable -Data $data -TableName 'Sales' }</code>
///   <para>Writes two rows and formats them as a styled Excel table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTable", DefaultParameterSetName = ParameterSetObjects)]
[Alias("ExcelTable")]
public sealed class AddOfficeExcelTableCommand : PSCmdlet
{
    private const string ParameterSetObjects = "Objects";
    private const string ParameterSetDataTable = "DataTable";

    /// <summary>Source objects to convert into table rows.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetObjects)]
    public object[] Data { get; set; } = Array.Empty<object>();

    /// <summary>Source <see cref="DataTable"/> to write directly.</summary>
    [Parameter(Mandatory = true, ParameterSetName = ParameterSetDataTable, ValueFromPipeline = true)]
    public DataTable? DataTable { get; set; }

    /// <summary>Starting row for the data (1-based).</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Starting column for the data (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Skip writing headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Name to assign to the table.</summary>
    [Parameter]
    public string? TableName { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public string TableStyle { get; set; } = "TableStyleMedium9";

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
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        var table = ExcelTabularInputService.ToDataTable(GetSourceInput(), TableName);
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

        var range = sheet.InsertDataTableAsTable(
            table,
            startRow: StartRow,
            startColumn: StartColumn,
            includeHeaders: !NoHeader.IsPresent,
            tableName: ResolveTableName(table),
            style: style,
            includeAutoFilter: !NoAutoFilter.IsPresent);

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

    private IEnumerable<object?> GetSourceInput()
    {
        if (ParameterSetName == ParameterSetDataTable)
        {
            if (DataTable == null)
            {
                throw new PSArgumentNullException(nameof(DataTable));
            }

            return new object?[] { DataTable };
        }

        if (Data == null || Data.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data row.", nameof(Data));
        }

        return Data;
    }
}
