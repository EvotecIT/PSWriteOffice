using System;
using System.Collections.Generic;

namespace PSWriteOffice.Services.Table;

internal sealed class OfficeTableSpec
{
    public OfficeTableSpec(IReadOnlyList<IReadOnlyList<OfficeTableCellSpec>> rows)
    {
        Rows = rows ?? throw new ArgumentNullException(nameof(rows));
        Placements = OfficeTableGridPlanner.Plan(rows, out var columnCount);
        ColumnCount = columnCount;
    }

    public IReadOnlyList<IReadOnlyList<OfficeTableCellSpec>> Rows { get; }

    public IReadOnlyList<OfficeTableCellPlacement> Placements { get; }

    public int RowCount => Rows.Count;

    public int ColumnCount { get; }
}

internal sealed class OfficeTableCellPlacement
{
    public OfficeTableCellPlacement(int rowIndex, int columnIndex, OfficeTableCellSpec cell)
    {
        RowIndex = rowIndex;
        ColumnIndex = columnIndex;
        Cell = cell ?? throw new ArgumentNullException(nameof(cell));
    }

    public int RowIndex { get; }

    public int ColumnIndex { get; }

    public OfficeTableCellSpec Cell { get; }
}
