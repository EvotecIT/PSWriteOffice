using System;
using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Get, "OfficeExcelWorkSheet")]
public class GetOfficeExcelWorkSheetCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    [Alias("Excel", "ExcelDocument")]
    public XLWorkbook Workbook { get; set; } = null!;

    [Parameter]
    [Alias("Name")]
    public string? WorksheetName { get; set; }

    [Parameter]
    public int? Index { get; set; }

    [Parameter]
    public SwitchParameter NameOnly { get; set; }

    protected override void ProcessRecord()
    {
        if (!string.IsNullOrEmpty(WorksheetName))
        {
            var worksheet = ExcelDocumentService.GetWorksheet(Workbook, WorksheetName);
            if (worksheet == null)
            {
                WriteError(new ErrorRecord(new ArgumentException($"Worksheet '{WorksheetName}' not found."),
                    "WorksheetNotFound", ErrorCategory.ObjectNotFound, WorksheetName));
                return;
            }
            WriteWorksheet(worksheet);
            return;
        }

        if (Index.HasValue)
        {
            var worksheet = ExcelDocumentService.GetWorksheet(Workbook, Index.Value);
            if (worksheet == null)
            {
                WriteError(new ErrorRecord(new ArgumentException($"Worksheet with index {Index.Value} not found."),
                    "WorksheetNotFound", ErrorCategory.ObjectNotFound, Index));
                return;
            }
            WriteWorksheet(worksheet);
            return;
        }

        foreach (var worksheet in ExcelDocumentService.GetWorksheets(Workbook))
        {
            WriteWorksheet(worksheet);
        }
    }

    private void WriteWorksheet(IXLWorksheet worksheet)
    {
        if (NameOnly.IsPresent)
        {
            WriteObject(worksheet.Name);
        }
        else
        {
            WriteObject(worksheet);
        }
    }
}
