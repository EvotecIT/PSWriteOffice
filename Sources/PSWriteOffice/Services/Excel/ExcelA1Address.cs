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

    internal static bool TryGetRangeStartRow(string? range, out int row)
    {
        row = 0;
        if (string.IsNullOrWhiteSpace(range))
        {
            return false;
        }

        var rangeValue = (range ?? string.Empty).Trim();
        var separatorIndex = rangeValue.IndexOf(':');
        var startReference = separatorIndex >= 0 ? rangeValue.Substring(0, separatorIndex) : rangeValue;
        var digitStart = 0;
        while (digitStart < startReference.Length && !char.IsDigit(startReference[digitStart]))
        {
            digitStart++;
        }

        if (digitStart >= startReference.Length)
        {
            return false;
        }

        return int.TryParse(startReference.Substring(digitStart), NumberStyles.None, CultureInfo.InvariantCulture, out row);
    }
}
