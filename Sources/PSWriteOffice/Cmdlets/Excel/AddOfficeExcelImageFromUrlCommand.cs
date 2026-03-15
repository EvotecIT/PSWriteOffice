using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Adds an image from a URL anchored to a worksheet cell.</summary>
/// <example>
///   <summary>Insert an image from a URL at B2.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Add-OfficeExcelImageFromUrl -Address 'B2' -Url 'https://example.org/logo.png' -WidthPixels 120 -HeightPixels 40 }</code>
///   <para>Downloads the remote image and anchors it to cell B2.</para>
/// </example>
[Cmdlet(VerbsCommon.Add, "OfficeExcelImageFromUrl", DefaultParameterSetName = ParameterSetContext)]
[Alias("ExcelImageFromUrl")]
public sealed class AddOfficeExcelImageFromUrlCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Workbook to operate on outside the DSL context.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public ExcelDocument Document { get; set; } = null!;

    /// <summary>Worksheet name when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public string? Sheet { get; set; }

    /// <summary>Worksheet index (0-based) when using <see cref="Document"/>.</summary>
    [Parameter(ParameterSetName = ParameterSetDocument)]
    public int? SheetIndex { get; set; }

    /// <summary>Image URL to download.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public string Url { get; set; } = string.Empty;

    /// <summary>1-based row index.</summary>
    [Parameter]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter]
    public string? Address { get; set; }

    /// <summary>Image width in pixels.</summary>
    [Parameter]
    public int WidthPixels { get; set; } = 96;

    /// <summary>Image height in pixels.</summary>
    [Parameter]
    public int HeightPixels { get; set; } = 32;

    /// <summary>Horizontal offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetXPixels { get; set; }

    /// <summary>Vertical offset in pixels from the cell origin.</summary>
    [Parameter]
    public int OffsetYPixels { get; set; }

    /// <summary>Emit the worksheet after inserting the image.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (WidthPixels <= 0 || HeightPixels <= 0)
        {
            throw new PSArgumentException("WidthPixels and HeightPixels must be greater than zero.");
        }

        var sheet = ResolveSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        sheet.AddImageFromUrlAt(row, column, Url, WidthPixels, HeightPixels, OffsetXPixels, OffsetYPixels);

        if (PassThru.IsPresent)
        {
            WriteObject(sheet);
        }
    }

    private ExcelSheet ResolveSheet()
    {
        if (ParameterSetName == ParameterSetDocument)
        {
            if (Document == null)
            {
                throw new PSArgumentException("Provide an Excel document.");
            }

            return ExcelSheetResolver.Resolve(Document, Sheet, SheetIndex);
        }

        var context = ExcelDslContext.Require(this);
        return context.RequireSheet();
    }
}
