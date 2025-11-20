using System;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes tabular data to the current worksheet and formats it as an Excel table.</summary>
/// <para>Accepts objects (PSCustomObject, hashtables, POCOs) and converts them into an Excel table with optional styling.</para>
/// <example>
///   <summary>Insert a table starting at A1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$data = @([pscustomobject]@{ Region='NA'; Revenue=100 }, [pscustomobject]@{ Region='EMEA'; Revenue=150 })
///   ExcelSheet 'Data' { Add-OfficeExcelTable -Data $data -TableName 'Sales' }</code>
///   <para>Writes two rows and formats them as a styled Excel table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelTable")]
[Alias("ExcelTable")]
public sealed class AddOfficeExcelTableCommand : PSCmdlet
{
    /// <summary>Source objects to convert into table rows.</summary>
    [Parameter(Mandatory = true)]
    public object[] Data { get; set; } = Array.Empty<object>();

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

    /// <summary>Return the created range string.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        if (Data == null || Data.Length == 0)
        {
            throw new PSArgumentException("Provide at least one data row.", nameof(Data));
        }

        var table = ExcelDataTableBuilder.FromObjects(Data);
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
            tableName: TableName,
            style: style,
            includeAutoFilter: !NoAutoFilter.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(range);
        }
    }
}
