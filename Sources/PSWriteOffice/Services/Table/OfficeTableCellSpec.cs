using System;

namespace PSWriteOffice.Services.Table;

/// <summary>Describes a logical table cell that can be rendered by multiple Office table surfaces.</summary>
public sealed class OfficeTableCellSpec
{
    /// <summary>Creates a logical table cell.</summary>
    /// <param name="text">Cell text.</param>
    /// <param name="columnSpan">Number of logical columns covered by the cell.</param>
    /// <param name="rowSpan">Number of logical rows covered by the cell.</param>
    public OfficeTableCellSpec(string? text, int columnSpan = 1, int rowSpan = 1)
    {
        if (columnSpan < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(columnSpan), "Column span must be at least 1.");
        }

        if (rowSpan < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(rowSpan), "Row span must be at least 1.");
        }

        Text = text ?? string.Empty;
        ColumnSpan = columnSpan;
        RowSpan = rowSpan;
    }

    /// <summary>Cell text.</summary>
    public string Text { get; }

    /// <summary>Number of logical columns covered by the cell.</summary>
    public int ColumnSpan { get; }

    /// <summary>Number of logical rows covered by the cell.</summary>
    public int RowSpan { get; }

    internal bool HasSpan => ColumnSpan > 1 || RowSpan > 1;
}
