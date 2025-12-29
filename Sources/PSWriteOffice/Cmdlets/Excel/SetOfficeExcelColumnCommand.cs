using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes values or formatting to a column in the current worksheet.</summary>
/// <example>
///   <summary>Populate a column and auto-fit it.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelColumn -Column 1 -Values 'North','South' -AutoFit }</code>
///   <para>Writes values into column A and adjusts the width.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelColumn")]
[Alias("ExcelColumn")]
public sealed class SetOfficeExcelColumnCommand : PSCmdlet
{
    /// <summary>1-based column index.</summary>
    [Parameter(Position = 0)]
    public int? Column { get; set; }

    /// <summary>Column letter reference (e.g., A, BC).</summary>
    [Parameter]
    [Alias("ColumnLetter", "Letter")]
    public string? ColumnName { get; set; }

    /// <summary>Values to write down the column.</summary>
    [Parameter]
    public object[]? Values { get; set; }

    /// <summary>Starting row index (1-based) for values.</summary>
    [Parameter]
    public int StartRow { get; set; } = 1;

    /// <summary>Column width to apply.</summary>
    [Parameter]
    public double? Width { get; set; }

    /// <summary>Hide or show the column.</summary>
    [Parameter]
    public bool? Hidden { get; set; }

    /// <summary>Auto-fit the column width after updates.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        var columnIndex = ExcelHostExtensions.ResolveColumnIndex(Column, ColumnName);
        if (columnIndex < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Column), "Column index must be 1 or greater.");
        }

        var hasAction = false;

        if (Values != null && Values.Length > 0)
        {
            if (StartRow < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(StartRow), "StartRow must be 1 or greater.");
            }

            var cells = new List<(int Row, int Column, object Value)>(Values.Length);
            for (int i = 0; i < Values.Length; i++)
            {
                var value = Values[i] ?? string.Empty;
                cells.Add((StartRow + i, columnIndex, value));
            }
            sheet.CellValues(cells);
            hasAction = true;
        }

        if (Width.HasValue)
        {
            sheet.SetColumnWidth(columnIndex, Width.Value);
            hasAction = true;
        }

        if (Hidden.HasValue)
        {
            sheet.SetColumnHidden(columnIndex, Hidden.Value);
            hasAction = true;
        }

        if (AutoFit.IsPresent)
        {
            sheet.AutoFitColumn(columnIndex);
            hasAction = true;
        }

        if (!hasAction)
        {
            throw new PSArgumentException("Provide -Values, -Width, -Hidden, or -AutoFit to update the column.");
        }
    }
}
