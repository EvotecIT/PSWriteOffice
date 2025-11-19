using System;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Text.RegularExpressions;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelHostExtensions
{
    private static readonly Regex AddressRegex = new(@"^([A-Za-z]+)(\d+)$", RegexOptions.Compiled);

    public static ExcelSheet GetOrCreateSheet(this ExcelDocument document, string? name, SheetNameValidationMode validationMode)
    {
        if (document == null) throw new ArgumentNullException(nameof(document));

        if (!string.IsNullOrWhiteSpace(name))
        {
            var existing = document.Sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            return document.AddWorkSheet(name ?? string.Empty, validationMode);
        }

        var sheets = document.Sheets;
        if (sheets == null || sheets.Count == 0)
        {
            return document.AddWorkSheet(string.Empty, SheetNameValidationMode.None);
        }

        return sheets![sheets.Count - 1];
    }

    public static (int Row, int Column) ResolveCellAddress(int? row, int? column, string? address)
    {
        if (!string.IsNullOrWhiteSpace(address))
        {
            var match = AddressRegex.Match(address.Trim());
            if (!match.Success)
            {
                throw new ArgumentException($"Address '{address}' is not a valid A1 reference.", nameof(address));
            }

            string columnPart = match.Groups[1].Value.ToUpperInvariant();
            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var rowIndex))
            {
                throw new ArgumentException($"Address '{address}' is not a valid A1 reference.", nameof(address));
            }

            return (rowIndex, ColumnLettersToIndex(columnPart));
        }

        if (!row.HasValue || !column.HasValue)
        {
            throw new ArgumentException("Specify either -Address or both -Row and -Column.");
        }

        return (row.Value, column.Value);
    }

    private static int ColumnLettersToIndex(string letters)
    {
        int result = 0;
        foreach (char c in letters)
        {
            if (c < 'A' || c > 'Z') throw new ArgumentException($"Invalid column letter '{c}'.", nameof(letters));
            result = result * 26 + (c - 'A' + 1);
        }
        return result;
    }
}
