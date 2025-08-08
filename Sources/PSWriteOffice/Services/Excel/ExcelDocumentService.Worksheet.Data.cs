using System;
using System.Collections.Generic;
using System.Globalization;
using ClosedXML.Excel;

namespace PSWriteOffice.Services.Excel;

public static partial class ExcelDocumentService
{
    public static IEnumerable<IDictionary<string, object?>> GetWorksheetData(IXLWorksheet worksheet, CultureInfo? culture = null)
    {
        var headers = new List<string>();
        var range = worksheet.RangeUsed();
        if (range == null)
        {
            yield break;
        }

        foreach (var cell in range.Row(1).Cells())
        {
            var text = cell.GetString();
            var name = text != string.Empty ? text : $"NoName{cell.Address}";
            if (headers.Contains(name))
            {
                name += cell.Address.ToString();
            }
            headers.Add(name);
        }

        var lastRow = range.RowCount();
        foreach (var row in range.Rows(2, lastRow))
        {
            var rowData = new Dictionary<string, object?>();
            for (var i = 0; i < headers.Count; i++)
            {
                var cell = row.Cell(i + 1);
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
