using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Set, "OfficeExcelWorkSheetStyle", DefaultParameterSetName = "ByName")]
public class SetOfficeExcelWorkSheetStyleCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public XLWorkbook Excel { get; set; } = null!;

    [Parameter(ParameterSetName = "ByObject", Mandatory = true)]
    public IXLWorksheet? Worksheet { get; set; }

    [Parameter(ParameterSetName = "ByName", Mandatory = true)]
    public string? WorksheetName { get; set; }

    [Parameter(ParameterSetName = "ByIndex", Mandatory = true)]
    public int? Index { get; set; }

    [Parameter]
    public string? TabColor { get; set; }

    protected override void ProcessRecord()
    {
        var worksheet = Worksheet;
        if (worksheet == null)
        {
            if (WorksheetName != null)
            {
                worksheet = ExcelDocumentService.GetWorksheet(Excel, WorksheetName);
            }
            else if (Index.HasValue)
            {
                worksheet = ExcelDocumentService.GetWorksheet(Excel, Index.Value);
            }
        }

        if (worksheet == null)
        {
            WriteWarning("Worksheet not found.");
            return;
        }

        var color = ColorService.GetColor(TabColor);
        if (color != null)
        {
            ExcelDocumentService.SetWorksheetTabColor(worksheet, color);
        }
    }
}
