using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Table;

internal static class TableViewProjection
{
    public static object[] Project(IReadOnlyList<object> rows, OfficeTableView view)
    {
        return view switch
        {
            OfficeTableView.Normal => rows.ToArray(),
            OfficeTableView.Transpose => TransposeRows(ExpandSingleTabularInput(rows)),
            _ => throw new PSArgumentException($"Unsupported table view '{view}'.", nameof(view))
        };
    }

    private static IReadOnlyList<object> ExpandSingleTabularInput(IReadOnlyList<object> rows)
    {
        if (rows.Count != 1)
        {
            return rows;
        }

        return rows[0] switch
        {
            DataTable table => table.Rows.Cast<DataRow>().Cast<object>().ToArray(),
            DataView view => view.Cast<DataRowView>().Cast<object>().ToArray(),
            _ => rows
        };
    }

    private static object[] TransposeRows(IReadOnlyList<object> rows)
    {
        var maps = rows.Select(BuildPropertyMap).ToList();
        var columnNames = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var map in maps)
        {
            foreach (var key in map.Keys)
            {
                if (seen.Add(key))
                {
                    columnNames.Add(key);
                }
            }
        }

        if (columnNames.Count == 0)
        {
            return Array.Empty<object>();
        }

        var transposed = new object[columnNames.Count];
        for (var columnIndex = 0; columnIndex < columnNames.Count; columnIndex++)
        {
            var propertyName = columnNames[columnIndex];
            var transposedRow = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Property"] = propertyName
            };

            for (var rowIndex = 0; rowIndex < maps.Count; rowIndex++)
            {
                var header = "Row" + (rowIndex + 1).ToString(CultureInfo.InvariantCulture);
                transposedRow[header] = maps[rowIndex].TryGetValue(propertyName, out var value)
                    ? value
                    : null;
            }

            transposed[columnIndex] = transposedRow;
        }

        return transposed;
    }

    private static Dictionary<string, object?> BuildPropertyMap(object row)
    {
        if (row is DataRow dataRow)
        {
            return BuildDataRowMap(dataRow);
        }

        if (row is DataRowView rowView)
        {
            return BuildDataRowMap(rowView.Row);
        }

        var normalized = PowerShellObjectNormalizer.NormalizeItem(row);
        if (normalized is IDictionary dictionary)
        {
            var mapped = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(key) || mapped.ContainsKey(key!))
                {
                    continue;
                }

                mapped[key!] = entry.Value;
            }

            if (mapped.Count > 0)
            {
                return mapped;
            }
        }

        var psObject = PSObject.AsPSObject(normalized);
        var properties = psObject.Properties
            .Where(property => property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property)
            .Where(property => !string.IsNullOrWhiteSpace(property.Name))
            .ToList();

        if (properties.Count == 0)
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["Value"] = normalized
            };
        }

        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            if (!result.ContainsKey(property.Name))
            {
                result[property.Name] = property.Value;
            }
        }

        return result;
    }

    private static Dictionary<string, object?> BuildDataRowMap(DataRow row)
    {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (DataColumn column in row.Table.Columns)
        {
            result[column.ColumnName] = row.IsNull(column) ? null : row[column];
        }

        return result;
    }
}
