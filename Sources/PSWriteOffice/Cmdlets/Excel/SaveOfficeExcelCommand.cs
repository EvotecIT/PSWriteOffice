using System;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Save, "OfficeExcel", SupportsShouldProcess = true)]
public class SaveOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public XLWorkbook Workbook { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public SwitchParameter Show { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            if (ShouldProcess(FilePath, "Save workbook"))
            {
                ExcelDocumentService.SaveWorkbook(Workbook, FilePath, Show);
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelSaveFailed", ErrorCategory.WriteError, FilePath));
        }
    }
}
