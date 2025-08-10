using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.New, "OfficeExcelValue")]
public class NewOfficeExcelValueCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public IXLWorksheet Worksheet { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public object? Value { get; set; }

    [Parameter(Mandatory = true)]
    public int Row { get; set; }

    [Parameter(Mandatory = true)]
    public int Column { get; set; }

    [Parameter]
    public string? DateFormat { get; set; }

    [Parameter]
    public string? NumberFormat { get; set; }

    [Parameter]
    public int? FormatID { get; set; }

    protected override void ProcessRecord()
    {
        var cell = ExcelDocumentService.SetCellValue(Worksheet, Row, Column, Value, DateFormat, NumberFormat, FormatID);
        WriteObject(cell);
    }
}
