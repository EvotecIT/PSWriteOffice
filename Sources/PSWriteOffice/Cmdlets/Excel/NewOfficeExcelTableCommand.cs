using System;
using System.Collections.Generic;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.New, "OfficeExcelTable")]
public class NewOfficeExcelTableCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public IXLWorksheet Worksheet { get; set; } = null!;

    [Parameter(Mandatory = true, ValueFromPipeline = true)]
    [Alias("Data")] 
    public PSObject[] DataTable { get; set; } = Array.Empty<PSObject>();

    [Parameter]
    [Alias("Row")]
    public int StartRow { get; set; } = 1;

    [Parameter]
    [Alias("Column")]
    public int StartColumn { get; set; } = 1;

    [Parameter]
    public SwitchParameter ShowRowStripes { get; set; }

    [Parameter]
    public SwitchParameter ShowColumnStripes { get; set; }

    [Parameter]
    public SwitchParameter DisableAutoFilter { get; set; }

    [Parameter]
    public SwitchParameter HideHeaderRow { get; set; }

    [Parameter]
    public SwitchParameter ShowTotalsRow { get; set; }

    [Parameter]
    public SwitchParameter EmphasizeFirstColumn { get; set; }

    [Parameter]
    public SwitchParameter EmphasizeLastColumn { get; set; }

    [Parameter]
    public XLTableTheme Theme { get; set; } = XLTableTheme.None;

    [Parameter]
    public XLTransposeOptions? Transpose { get; set; }

    private readonly List<PSObject> _data = new();

    protected override void ProcessRecord()
    {
        if (DataTable != null)
        {
            foreach (var item in DataTable)
            {
                _data.Add(item);
            }
        }
    }

    protected override void EndProcessing()
    {
        var tableData = new List<IDictionary<string, object?>>();
        foreach (var item in _data)
        {
            if (item.BaseObject is IDictionary<string, object?> dict)
            {
                tableData.Add(dict);
            }
            else
            {
                var result = new Dictionary<string, object?>();
                foreach (var prop in item.Properties)
                {
                    result[prop.Name] = prop.Value;
                }
                tableData.Add(result);
            }
        }

        var table = ExcelDocumentService.InsertTable(
            Worksheet,
            tableData,
            StartRow,
            StartColumn,
            Theme,
            ShowRowStripes,
            ShowColumnStripes,
            !DisableAutoFilter,
            !HideHeaderRow,
            ShowTotalsRow,
            EmphasizeFirstColumn,
            EmphasizeLastColumn,
            Transpose);

        WriteObject(table);
    }
}
