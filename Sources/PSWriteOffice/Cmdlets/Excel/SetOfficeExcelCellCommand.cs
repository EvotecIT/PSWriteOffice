using System.Management.Automation;
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

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();
        var (row, column) = ExcelHostExtensions.ResolveCellAddress(Row, Column, Address);

        if (Value != null || Formula != null || NumberFormat != null)
        {
            sheet.Cell(row, column, Value, Formula, NumberFormat);
        }
        else
        {
            throw new PSArgumentException("Provide -Value, -Formula, or -NumberFormat to modify the cell.");
        }
    }
}
