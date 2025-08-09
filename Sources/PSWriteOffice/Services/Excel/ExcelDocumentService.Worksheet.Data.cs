using System;
using System.Collections.Generic;
using System.Globalization;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IEnumerable<IDictionary<string, object?>> GetWorksheetData(IXLWorksheet worksheet, CultureInfo? culture = null)
        => GetWorksheetData(worksheet, null, null, null, null, null, false, culture);

    public static IEnumerable<IDictionary<string, object?>> GetWorksheetData(
        IXLWorksheet worksheet,
        int? startRow,
        int? endRow,
        int? startColumn,
        int? endColumn,
        int? headerRow,
        bool noHeader,
        CultureInfo? culture = null)
    {
        var range = worksheet.RangeUsed();
        if (range == null)
        {
            yield break;
        }

        var firstRow = startRow ?? range.FirstRow().RowNumber();
        var lastRow = endRow ?? range.LastRow().RowNumber();
        var firstColumn = startColumn ?? range.FirstColumn().ColumnNumber();
        var lastColumn = endColumn ?? range.LastColumn().ColumnNumber();

        var headers = new List<string>();
        if (!noHeader)
        {
            var headerRowNumber = headerRow ?? firstRow;
            for (var col = firstColumn; col <= lastColumn; col++)
            {
                var cell = worksheet.Cell(headerRowNumber, col);
                var text = cell.GetString();
                var name = text != string.Empty ? text : $"NoName{cell.Address}";
                if (headers.Contains(name))
                {
                    name += cell.Address.ToString();
                }
                headers.Add(name);
            }
        }
        else
        {
            for (var col = firstColumn; col <= lastColumn; col++)
            {
                headers.Add($"Column{col - firstColumn + 1}");
            }
        }

        var headerRowActual = headerRow ?? firstRow;
        for (var rowNumber = firstRow; rowNumber <= lastRow; rowNumber++)
        {
            if (!noHeader && rowNumber == headerRowActual)
            {
                continue;
            }

            var rowData = new Dictionary<string, object?>();
            for (var i = 0; i < headers.Count; i++)
            {
                var cell = worksheet.Cell(rowNumber, firstColumn + i);
                object? value = cell.CachedValue;

                if (culture != null && value is string textValue)
                {
                    if (DateTime.TryParse(textValue, culture, DateTimeStyles.None, out var date))
                    {
                        value = date;
                    }
                    else if (double.TryParse(textValue, NumberStyles.Any, culture, out var number))
                    {
                        value = number;
                    }
                }

                rowData[headers[i]] = value;
            }

            yield return rowData;
        }
    }
}
