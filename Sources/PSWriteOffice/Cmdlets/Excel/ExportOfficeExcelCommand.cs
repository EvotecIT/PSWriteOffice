using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Export, "OfficeExcel")]
public class ExportOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    [Alias("Name")]
    public string WorksheetName { get; set; } = "Sheet1";

    [Parameter(ValueFromPipeline = true)]
    [Alias("TargetData")]
    public PSObject[] DataTable { get; set; } = Array.Empty<PSObject>();

    [Parameter]
    public int Row { get; set; } = 1;

    [Parameter]
    public int Column { get; set; } = 1;

    [Parameter]
    public SwitchParameter Show { get; set; }

    [Parameter]
    public SwitchParameter AllProperties { get; set; }

    [Parameter]
    public XLTransposeOptions? Transpose { get; set; }

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
        try
        {
            var workbook = File.Exists(FilePath)
                ? ExcelDocumentService.LoadWorkbook(FilePath)
                : ExcelDocumentService.CreateWorkbook();

            var worksheet = ExcelDocumentService.AddWorksheet(workbook, WorksheetName, WorksheetExistOption.Replace);

            var tableData = _data.Select(item =>
            {
                if (item.BaseObject is IDictionary<string, object?> dict)
                {
                    return dict;
                }

                var psobj = item;
                var result = new Dictionary<string, object?>();
                foreach (var prop in psobj.Properties)
                {
                    result[prop.Name] = prop.Value;
                }
                return (IDictionary<string, object?>)result;
            }).ToList();

            ExcelDocumentService.InsertTable(
                worksheet,
                tableData,
                Row,
                Column,
                Theme,
                ShowRowStripes,
                ShowColumnStripes,
                !DisableAutoFilter,
                !HideHeaderRow,
                ShowTotalsRow,
                EmphasizeFirstColumn,
                EmphasizeLastColumn,
                Transpose);

            ExcelDocumentService.SaveWorkbook(workbook, FilePath, Show);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExportOfficeExcelFailed", ErrorCategory.NotSpecified, FilePath));
        }
    }
}
