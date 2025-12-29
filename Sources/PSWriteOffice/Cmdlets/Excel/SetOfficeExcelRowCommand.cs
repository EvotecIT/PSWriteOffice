using System;
using System.Collections.Generic;
using System.Management.Automation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Writes a row of values to the current worksheet.</summary>
/// <example>
///   <summary>Write a row starting at column A.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ExcelSheet 'Data' { Set-OfficeExcelRow -Row 2 -Values 'North', 1200 }</code>
///   <para>Writes two values into row 2, columns A and B.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficeExcelRow")]
[Alias("ExcelRow")]
public sealed class SetOfficeExcelRowCommand : PSCmdlet
{
    /// <summary>1-based row index.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    public int Row { get; set; }

    /// <summary>Values to write across the row.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public object[] Values { get; set; } = Array.Empty<object>();

    /// <summary>Starting column index (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var context = ExcelDslContext.Require(this);
        var sheet = context.RequireSheet();

        if (Row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(Row), "Row index must be 1 or greater.");
        }

        if (StartColumn < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(StartColumn), "StartColumn must be 1 or greater.");
        }

        if (Values.Length == 0)
        {
            throw new PSArgumentException("Provide at least one value.", nameof(Values));
        }

        var cells = new List<(int Row, int Column, object Value)>(Values.Length);
        for (int i = 0; i < Values.Length; i++)
        {
            var value = Values[i] ?? string.Empty;
            cells.Add((Row, StartColumn + i, value));
        }

        sheet.CellValues(cells);
    }
}
