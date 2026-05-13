using System;
using System.Globalization;
using System.Text;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelA1Address
{
    internal static string CellReference(int row, int column)
    {
        if (row < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(row), "Row index must be 1 or greater.");
        }

        if (column < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(column), "Column index must be 1 or greater.");
        }

        var builder = new StringBuilder();
        var value = column;
        while (value > 0)
        {
            value--;
            builder.Insert(0, (char)('A' + value % 26));
            value /= 26;
        }

        builder.Append(row.ToString(CultureInfo.InvariantCulture));
        return builder.ToString();
    }
}
