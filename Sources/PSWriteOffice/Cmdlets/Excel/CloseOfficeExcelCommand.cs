using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Close, "OfficeExcel", SupportsShouldProcess = true)]
public class CloseOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public XLWorkbook Workbook { get; set; } = null!;

    protected override void ProcessRecord()
    {
        if (ShouldProcess("Workbook", "Close workbook"))
        {
            ExcelDocumentService.CloseWorkbook(Workbook);
        }
    }
}
