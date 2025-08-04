using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IXLCell SetCellValue(IXLWorksheet worksheet, int row, int column, object? value, string? dateFormat = null,
        string? numberFormat = null, int? formatId = null)
    {
        var cell = worksheet.Cell(row, column);
        if (value == null)
        {
            cell.Value = string.Empty;
        }
        else
        {
            try
            {
                cell.Value = XLCellValue.FromObject(value);
            }
            catch
            {
                cell.Value = value.ToString();
            }
        }

        if (!string.IsNullOrEmpty(dateFormat))
        {
            cell.Style.DateFormat.Format = dateFormat;
        }

        if (!string.IsNullOrEmpty(numberFormat))
        {
            cell.Style.NumberFormat.Format = numberFormat;
        }

        if (formatId.HasValue)
        {
            cell.Style.NumberFormat.NumberFormatId = formatId.Value;
        }

        return cell;
    }
}
