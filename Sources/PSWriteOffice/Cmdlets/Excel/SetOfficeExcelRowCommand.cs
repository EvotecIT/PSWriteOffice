using System;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO.Excel;
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
    [Parameter(Position = 1)]
    public object[]? Values { get; set; }

    /// <summary>Starting column index (1-based).</summary>
    [Parameter]
    public int StartColumn { get; set; } = 1;

    /// <summary>Explicit row height in points.</summary>
    [Parameter]
    public double? Height { get; set; }

    /// <summary>Clear any custom row height.</summary>
    [Parameter]
    public SwitchParameter ClearHeight { get; set; }

    /// <summary>Auto-fit row height after writing values and applying style.</summary>
    [Parameter]
    public SwitchParameter AutoFit { get; set; }

    /// <summary>Hide or unhide the row. Use -Hidden:$false to unhide.</summary>
    [Parameter]
    public bool? Hidden { get; set; }

    /// <summary>Apply or clear bold styling across the target row span.</summary>
    [Parameter]
    public bool? Bold { get; set; }

    /// <summary>Apply or clear italic styling across the target row span.</summary>
    [Parameter]
    public bool? Italic { get; set; }

    /// <summary>Apply or clear underline styling across the target row span.</summary>
    [Parameter]
    public bool? Underline { get; set; }

    /// <summary>Apply or clear wrap text across the target row span.</summary>
    [Parameter]
    public bool? WrapText { get; set; }

    /// <summary>Apply a font family across the target row span.</summary>
    [Parameter]
    public string? FontName { get; set; }

    /// <summary>Apply a background color across the target row span.</summary>
    [Parameter]
    public string? BackgroundColor { get; set; }

    /// <summary>First 1-based column affected by style options.</summary>
    [Parameter]
    public int? FirstColumn { get; set; }

    /// <summary>Last 1-based column affected by style options.</summary>
    [Parameter]
    public int? LastColumn { get; set; }

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

        var values = Values ?? Array.Empty<object>();
        var hasLayout = HasLayoutOptions();
        if (values.Length == 0 && !hasLayout)
        {
            throw new PSArgumentException("Provide row values or at least one layout/style option.", nameof(Values));
        }

        if (values.Length > 0)
        {
            var cells = new List<(int Row, int Column, object Value)>(values.Length);
            for (int i = 0; i < values.Length; i++)
            {
                var value = values[i] ?? string.Empty;
                cells.Add((Row, StartColumn + i, value));
            }

            sheet.CellValues(cells);
        }

        if (hasLayout)
        {
            sheet.SetRowLayout(Row, new ExcelRowLayoutOptions {
                Height = Height,
                ClearHeight = ClearHeight.IsPresent,
                AutoFit = AutoFit.IsPresent,
                Hidden = Hidden,
                Bold = Bold,
                Italic = Italic,
                Underline = Underline,
                WrapText = WrapText,
                FontName = FontName,
                BackgroundColor = BackgroundColor,
                FirstColumn = FirstColumn ?? (values.Length > 0 ? StartColumn : null),
                LastColumn = LastColumn ?? (values.Length > 0 ? StartColumn + values.Length - 1 : null)
            });
        }
    }

    private bool HasLayoutOptions()
    {
        return Height.HasValue ||
            ClearHeight.IsPresent ||
            AutoFit.IsPresent ||
            Hidden.HasValue ||
            Bold.HasValue ||
            Italic.HasValue ||
            Underline.HasValue ||
            WrapText.HasValue ||
            !string.IsNullOrWhiteSpace(FontName) ||
            !string.IsNullOrWhiteSpace(BackgroundColor) ||
            FirstColumn.HasValue ||
            LastColumn.HasValue;
    }
}
