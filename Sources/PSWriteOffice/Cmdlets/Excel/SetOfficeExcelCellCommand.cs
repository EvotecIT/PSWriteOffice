using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Sets a cell value, formula, or number format within the current worksheet.</summary>
/// <para>Supports A1 addresses or row/column coordinates for DSL-style composition.</para>
/// <example>
///   <summary>Write values to A1 and B1.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelCell -Address 'A1' -Value 'Region'; Set-OfficeExcelCell -Row 1 -Column 2 -Value 'Revenue' }</code>
///   <para>Writes two headers in the first row.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelCell")]
[Alias("ExcelCell")]
public sealed class SetOfficeExcelCellCommand : PSCmdlet
{
    /// <summary>1-based row index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Row { get; set; }

    /// <summary>1-based column index.</summary>
    [Parameter(ParameterSetName = "Coordinates")]
    public int? Column { get; set; }

    /// <summary>A1-style cell address (e.g., A1, C5).</summary>
    [Parameter(ParameterSetName = "Address")]
    public string? Address { get; set; }

    /// <summary>Cell value to assign.</summary>
    [Parameter]
    public object? Value { get; set; }

    /// <summary>Formula text (without leading =).</summary>
    [Parameter]
    public string? Formula { get; set; }

    /// <summary>Number format code to apply.</summary>
    [Parameter]
    public string? NumberFormat { get; set; }

    /// <summary>Solid background color as #RRGGBB or #AARRGGBB.</summary>
    [Parameter]
    public string? BackgroundColor { get; set; }

    /// <summary>Gradient start color as #RRGGBB or #AARRGGBB.</summary>
    [Parameter]
    public string? GradientFrom { get; set; }

    /// <summary>Gradient end color as #RRGGBB or #AARRGGBB.</summary>
    [Parameter]
    public string? GradientTo { get; set; }

    /// <summary>Linear gradient angle in degrees.</summary>
    [Parameter]
    public double GradientDegree { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);
        var hasValueChange = Value != null || Formula != null || NumberFormat != null;
        var hasStyleChange = !string.IsNullOrWhiteSpace(BackgroundColor)
            || !string.IsNullOrWhiteSpace(GradientFrom)
            || !string.IsNullOrWhiteSpace(GradientTo);

        if (!hasValueChange && !hasStyleChange)
        {
            throw new PSArgumentException("Provide -Value, -Formula, -NumberFormat, -BackgroundColor, or -GradientFrom/-GradientTo to modify the cell.");
        }

        if (!string.IsNullOrWhiteSpace(GradientFrom) ^ !string.IsNullOrWhiteSpace(GradientTo))
        {
            throw new PSArgumentException("Specify both -GradientFrom and -GradientTo for a gradient fill.");
        }

        if (hasValueChange)
        {
            sheet.Cell(row, column, Value, Formula, NumberFormat);
        }

        if (!string.IsNullOrWhiteSpace(BackgroundColor))
        {
            sheet.CellBackground(row, column, BackgroundColor!);
        }

        if (!string.IsNullOrWhiteSpace(GradientFrom) && !string.IsNullOrWhiteSpace(GradientTo))
        {
            sheet.CellGradientBackground(row, column, GradientFrom!, GradientTo!, GradientDegree);
        }
    }
}
