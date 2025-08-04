using System;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;
using ValidateScriptAttribute = PSWriteOffice.Validation.ValidateScriptAttribute;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Save, "OfficeExcel")]
public class SaveOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public XLWorkbook Workbook { get; set; } = null!;

    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    [ValidateScript("{ Test-Path $_ }")]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public SwitchParameter Show { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            ExcelDocumentService.SaveWorkbook(Workbook, FilePath, Show);
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelSaveFailed", ErrorCategory.WriteError, FilePath));
        }
    }
}
