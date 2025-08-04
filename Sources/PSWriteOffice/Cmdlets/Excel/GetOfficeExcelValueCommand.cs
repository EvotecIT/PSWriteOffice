using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Get, "OfficeExcelValue")]
public class GetOfficeExcelValueCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    [Alias("WorkSheet")]
    public IXLWorksheet Worksheet { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public int Row { get; set; }

    [Parameter(Mandatory = true)]
    public int Column { get; set; }

    protected override void ProcessRecord()
    {
        var cell = ExcelDocumentService.GetCell(Worksheet, Row, Column);
        WriteObject(cell);
    }
}
