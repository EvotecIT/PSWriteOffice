using ClosedXML.Excel;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Excel;

public class ExcelCellStyleOptions
{
    public string? Format { get; set; }
    public int? FormatId { get; set; }
    public XLColor? FontColor { get; set; }
    public XLColor? BackgroundColor { get; set; }
    public XLFillPatternValues? PatternType { get; set; }
    public bool? Bold { get; set; }
    public XLFontCharSet? FontCharSet { get; set; }
    public XLFontFamilyNumberingValues? FontFamilyNumbering { get; set; }
    public string? FontName { get; set; }
    public double? FontSize { get; set; }
    public bool? Italic { get; set; }
    public bool? Shadow { get; set; }
    public bool? Strikethrough { get; set; }
    public XLFontUnderlineValues? Underline { get; set; }
    public XLFontVerticalTextAlignmentValues? VerticalAlignment { get; set; }
}

public static partial class ExcelDocumentService
{
    public static void SetCellStyle(IXLWorksheet worksheet, int row, int column, ExcelCellStyleOptions options)
    {
        var cell = worksheet.Cell(row, column);

        if (!string.IsNullOrEmpty(options.Format))
        {
            cell.Style.NumberFormat.Format = options.Format;
        }
        else if (options.FormatId.HasValue)
        {
            cell.Style.NumberFormat.NumberFormatId = options.FormatId.Value;
        }

        if (options.FontColor != null)
        {
            cell.Style.Font.FontColor = options.FontColor;
        }

        if (options.Bold.HasValue)
        {
            cell.Style.Font.Bold = options.Bold.Value;
        }

        if (options.Italic.HasValue)
        {
            cell.Style.Font.Italic = options.Italic.Value;
        }

        if (options.Strikethrough.HasValue)
        {
            cell.Style.Font.Strikethrough = options.Strikethrough.Value;
        }

        if (options.Shadow.HasValue)
        {
            cell.Style.Font.Shadow = options.Shadow.Value;
        }

        if (options.FontSize.HasValue)
        {
            cell.Style.Font.FontSize = options.FontSize.Value;
        }

        if (options.Underline.HasValue)
        {
            cell.Style.Font.Underline = options.Underline.Value;
        }

        if (options.VerticalAlignment.HasValue)
        {
            cell.Style.Font.VerticalAlignment = options.VerticalAlignment.Value;
        }

        if (options.FontFamilyNumbering.HasValue)
        {
            cell.Style.Font.FontFamilyNumbering = options.FontFamilyNumbering.Value;
        }

        if (options.FontCharSet.HasValue)
        {
            cell.Style.Font.FontCharSet = options.FontCharSet.Value;
        }

        if (!string.IsNullOrEmpty(options.FontName))
        {
            cell.Style.Font.FontName = options.FontName;
        }

        if (options.BackgroundColor != null)
        {
            cell.Style.Fill.BackgroundColor = options.BackgroundColor;
        }

        if (options.PatternType.HasValue)
        {
            cell.Style.Fill.PatternType = options.PatternType.Value;
        }
    }
}
