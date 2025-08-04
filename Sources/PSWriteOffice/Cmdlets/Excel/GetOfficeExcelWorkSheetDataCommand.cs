using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Get, "OfficeExcelWorkSheetData")]
public class GetOfficeExcelWorkSheetDataCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    [Alias("WorkSheet")]
    public IXLWorksheet Worksheet { get; set; } = null!;

    protected override void ProcessRecord()
    {
        foreach (var row in ExcelDocumentService.GetWorksheetData(Worksheet))
        {
            var obj = new PSObject();
            foreach (var kv in row)
            {
                obj.Properties.Add(new PSNoteProperty(kv.Key, kv.Value));
            }
            WriteObject(obj);
        }
    }
}
