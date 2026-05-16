using System;
using System.Data;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes every table in a <see cref="DataSet"/> to separate Excel worksheets.</summary>
/// <para>Uses OfficeIMO.Excel DataSet ingestion so callers can provide data from any .NET provider without PSWriteOffice owning database connections.</para>
/// <example>
///   <summary>Insert each DataSet table as a worksheet.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficeExcel -Path .\report.xlsx { Add-OfficeExcelDataSet -DataSet $dataSet -AutoFit }</code>
///   <para>Creates one worksheet per DataTable and formats each range as an Excel table.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelDataSet")]
[Alias("ExcelDataSet")]
[OutputType(typeof(ExcelDataSetImportResult))]
public sealed class AddOfficeExcelDataSetCommand : PSCmdlet
{
    /// <summary>Source DataSet whose tables will become worksheets.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public DataSet? DataSet { get; set; }

    /// <summary>Write plain ranges instead of Excel tables.</summary>
    [Parameter]
    public SwitchParameter NoTable { get; set; }

    /// <summary>Skip writing headers.</summary>
    [Parameter]
    public SwitchParameter NoHeader { get; set; }

    /// <summary>Built-in table style to apply.</summary>
    [Parameter]
    public string TableStyle { get; set; } = "TableStyleMedium9";

    /// <summary>Disable AutoFilter dropdowns.</summary>
    [Parameter]
    public SwitchParameter NoAutoFilter { get; set; }

    /// <summary>Auto-fit imported table columns.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Return import metadata for each worksheet.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (DataSet == null)
        {
            throw new PSArgumentNullException(nameof(DataSet));
        }

        if (!Enum.TryParse(TableStyle, ignoreCase: true, out TableStyle style))
        {
            throw new PSArgumentException($"Unknown table style '{TableStyle}'.", nameof(TableStyle));
        }

        var context = ExcelDslContext.Require(this);
        var results = context.Document.InsertDataSet(
            DataSet,
            createTables: !NoTable.IsPresent,
            tableStyle: style,
            includeHeaders: !NoHeader.IsPresent,
            includeAutoFilter: !NoAutoFilter.IsPresent,
            autoFit: AutoFit.IsPresent);

        if (PassThru.IsPresent)
        {
            WriteObject(results, enumerateCollection: true);
        }
    }
}
