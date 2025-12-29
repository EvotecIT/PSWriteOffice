using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelHostExtensions
{

    public static ExcelSheet GetOrCreateSheet(this ExcelDocument document, string? name, SheetNameValidationMode validationMode)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        var sheetsCollection = document.Sheets ?? new List<ExcelSheet>();

        if (!string.IsNullOrWhiteSpace(name))
        {
            var existing = sheetsCollection.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            return document.AddWorkSheet(name ?? string.Empty, validationMode);
        }

        if (sheetsCollection.Count == 0)
        {
            return document.AddWorkSheet(string.Empty, SheetNameValidationMode.None);
        }

        return sheetsCollection[sheetsCollection.Count - 1];
    }

    public static (int Row, int Column) ResolveCellAddress(int? row, int? column, string? address)
    {
        if (!string.IsNullOrWhiteSpace(address))
        {
            var trimmedAddress = address!.Trim();
            var (rowIndex, columnIndex) = A1.ParseCellRef(trimmedAddress);
            if (rowIndex <= 0 || columnIndex <= 0)
            {
                throw new ArgumentException($"Address '{address}' is not a valid A1 reference.", nameof(address));
            }

            return (rowIndex, columnIndex);
        }

        if (!row.HasValue || !column.HasValue)
        {
            throw new ArgumentException("Specify either -Address or both -Row and -Column.");
        }

        return (row.Value, column.Value);
    }

    public static int ResolveColumnIndex(int? columnIndex, string? columnName)
    {
        if (!string.IsNullOrWhiteSpace(columnName))
        {
            var index = A1.ColumnLettersToIndex(columnName!.Trim());
            if (index <= 0)
            {
                throw new ArgumentException($"ColumnName '{columnName}' is not a valid column reference.", nameof(columnName));
            }
            return index;
        }

        if (!columnIndex.HasValue)
        {
            throw new ArgumentException("Specify either -Column or -ColumnName.");
        }

        return columnIndex.Value;
    }

}
