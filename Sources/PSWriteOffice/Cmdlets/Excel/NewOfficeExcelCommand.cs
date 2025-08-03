using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.New, "OfficeExcel")]
public class NewOfficeExcelCommand : PSCmdlet
{
    protected override void ProcessRecord()
    {
        var workbook = ExcelDocumentService.CreateWorkbook();
        WriteObject(workbook);
    }
}
