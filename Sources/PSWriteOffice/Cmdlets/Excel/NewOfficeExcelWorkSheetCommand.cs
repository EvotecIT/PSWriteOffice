using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.New, "OfficeExcelWorkSheet")]
public class NewOfficeExcelWorkSheetCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public XLWorkbook Workbook { get; set; } = null!;

    [Parameter(Mandatory = true)]
    [Alias("Name")]
    public string WorksheetName { get; set; } = string.Empty;

    [Parameter]
    [ValidateSet("Replace", "Skip", "Rename")]
    public string Option { get; set; } = "Skip";

    [Parameter]
    public string? TabColor { get; set; }

    protected override void ProcessRecord()
    {
        var option = Option switch
        {
            "Replace" => WorksheetExistOption.Replace,
            "Rename" => WorksheetExistOption.Rename,
            _ => WorksheetExistOption.Skip
        };

        XLColor? color = null;
        if (!string.IsNullOrEmpty(TabColor))
        {
            try
            {
                color = XLColor.FromHtml(TabColor);
            }
            catch
            {
                try
                {
                    color = XLColor.FromName(TabColor);
                }
                catch
                {
                    WriteWarning($"Tab color {TabColor} is not valid.");
                }
            }
        }

        var worksheet = ExcelDocumentService.AddWorksheet(Workbook, WorksheetName, option, color);
        WriteObject(worksheet);
    }
}
