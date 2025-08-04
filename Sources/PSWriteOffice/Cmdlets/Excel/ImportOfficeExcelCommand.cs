using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services.Excel;
using ValidateScriptAttribute = PSWriteOffice.Validation.ValidateScriptAttribute;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsData.Import, "OfficeExcel")]
public class ImportOfficeExcelCommand : PSCmdlet
{
    [Alias("LiteralPath")]
    [Parameter(Mandatory = true)]
    [ValidateNotNullOrEmpty]
    [ValidateScript("{ Test-Path $_ }")]
    public string FilePath { get; set; } = string.Empty;

    [Parameter]
    public string[]? WorkSheetName { get; set; }

    protected override void ProcessRecord()
    {
        try
        {
            var data = ExcelDocumentService.ImportWorkbook(FilePath, WorkSheetName);
            if (WorkSheetName != null && WorkSheetName.Length == 1 && data.TryGetValue(WorkSheetName[0], out var single))
            {
                WriteObject(single, true);
            }
            else if ((WorkSheetName == null || WorkSheetName.Length == 0) && data.Count == 1)
            {
                WriteObject(data.Values.First(), true);
            }
            else
            {
                WriteObject(data);
            }
        }
        catch (FileNotFoundException ex)
        {
            WriteError(new ErrorRecord(ex, "FileNotFound", ErrorCategory.ObjectNotFound, FilePath));
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ExcelImportFailed", ErrorCategory.InvalidOperation, FilePath));
        }
    }
}
