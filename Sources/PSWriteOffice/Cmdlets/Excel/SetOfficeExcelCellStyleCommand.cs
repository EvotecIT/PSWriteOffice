using System.Management.Automation;
using ClosedXML.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

[Cmdlet(VerbsCommon.Set, "OfficeExcelCellStyle")]
public class SetOfficeExcelCellStyleCommand : PSCmdlet
{
    [Parameter(Mandatory = true)]
    public IXLWorksheet Worksheet { get; set; } = null!;

    [Parameter(Mandatory = true)]
    public int Row { get; set; }

    [Parameter(Mandatory = true)]
    public int Column { get; set; }

    [Parameter]
    public string? Format { get; set; }

    [Parameter]
    public int? FormatID { get; set; }

    [Parameter]
    [Alias("Color")]
    public string? FontColor { get; set; }

    [Parameter]
    public string? BackGroundColor { get; set; }

    [Parameter]
    public bool? Bold { get; set; }

    [Parameter]
    public XLFontCharSet? FontCharSet { get; set; }

    [Parameter]
    public XLFontFamilyNumberingValues? FontFamilyNumbering { get; set; }

    [Parameter]
    public string? FontName { get; set; }

    [Parameter]
    public double? FontSize { get; set; }

    [Parameter]
    public bool? Italic { get; set; }

    [Parameter]
    public bool? Shadow { get; set; }

    [Parameter]
    public bool? Strikethrough { get; set; }

    [Parameter]
    public XLFontUnderlineValues? Underline { get; set; }

    [Parameter]
    public XLFontVerticalTextAlignmentValues? VerticalAlignment { get; set; }

    [Parameter]
    public XLFillPatternValues? PatternType { get; set; }

    protected override void ProcessRecord()
    {
        var options = new ExcelCellStyleOptions
        {
            Format = Format,
            FormatId = FormatID,
            FontColor = ColorService.GetColor(FontColor),
            BackgroundColor = ColorService.GetColor(BackGroundColor),
            Bold = Bold,
            FontCharSet = FontCharSet,
            FontFamilyNumbering = FontFamilyNumbering,
            FontName = FontName,
            FontSize = FontSize,
            Italic = Italic,
            Shadow = Shadow,
            Strikethrough = Strikethrough,
            Underline = Underline,
            VerticalAlignment = VerticalAlignment,
            PatternType = PatternType
        };

        ExcelDocumentService.SetCellStyle(Worksheet, Row, Column, options);
    }
}
