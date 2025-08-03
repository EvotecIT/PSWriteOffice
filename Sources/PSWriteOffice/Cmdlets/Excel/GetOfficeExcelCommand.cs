using System;
using System.IO;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Get, "OfficeExcel")]
public class GetOfficeExcelCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public string FilePath { get; set; } = string.Empty;

    protected override void ProcessRecord()
    {
        try
        {
            var workbook = ExcelDocumentService.LoadWorkbook(FilePath);
            WriteObject(workbook);
        }
        catch (FileNotFoundException ex)
        {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelLoadFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
