using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static void AddChart(string filePath, string worksheetName, string chartTitle)
    {
        using var document = SpreadsheetDocument.Open(filePath, true);
        var workbookPart = document.WorkbookPart;
        if (workbookPart == null)
        {
            return;
        }

        var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => s.Name == worksheetName);
        if (sheet == null)
        {
            return;
        }

        var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        var drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
        if (!worksheetPart.Worksheet.Elements<Drawing>().Any())
        {
            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            worksheetPart.Worksheet.Save();
        }

        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        var chartSpace = new C.ChartSpace();
        chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
        var chart = chartSpace.AppendChild(new C.Chart());

        if (!string.IsNullOrEmpty(chartTitle))
        {
            var title = new C.Title
            {
                ChartText = new C.ChartText
                {
                    RichText = new C.RichText(new A.Paragraph(new A.Run(new A.Text(chartTitle))))
                }
            };
            chart.AppendChild(title);
        }

        chart.AppendChild(new C.PlotArea(new C.Layout()));
        chartPart.ChartSpace = chartSpace;
        chartPart.ChartSpace.Save();
    }
}
