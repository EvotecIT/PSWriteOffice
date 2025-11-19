using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Management.Automation;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelDataTableBuilder
{
    public static DataTable FromObjects(IReadOnlyCollection<object?> items)
    {
        if (items == null || items.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(items));
        }

        var table = new DataTable("Data");
        var first = items.FirstOrDefault();
        if (first == null)
        {
            throw new ArgumentException("Data rows cannot be null.", nameof(items));
        }

        var columns = GetColumnNames(first);
        if (columns.Count == 0)
        {
            throw new InvalidOperationException("Unable to infer column names. Use objects with properties or dictionaries.");
        }

        foreach (var name in columns)
        {
            table.Columns.Add(name, typeof(object));
        }

        foreach (var item in items)
        {
            if (item == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }
            var row = table.NewRow();
            foreach (var column in columns)
            {
                row[column] = GetValue(item, column) ?? DBNull.Value;
            }
            table.Rows.Add(row);
        }

        return table;
    }

    private static IReadOnlyList<string> GetColumnNames(object item)
    {
        var ps = PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary dict)
        {
            var names = new List<string>();
            foreach (DictionaryEntry entry in dict)
            {
                if (entry.Key is string key && !string.IsNullOrWhiteSpace(key))
                {
                    names.Add(key);
                }
            }
            return names;
        }

        var noteProperties = ps.Properties
            .Where(p => p.MemberType == PSMemberTypes.NoteProperty || p.MemberType == PSMemberTypes.Property)
            .Select(p => p.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();
        if (noteProperties.Count > 0)
        {
            return noteProperties;
        }

        var reflectionProps = ps.BaseObject.GetType()
            .GetProperties()
            .Where(p => p.CanRead)
            .Select(p => p.Name)
            .ToList();

        return reflectionProps;
    }

    private static object? GetValue(object item, string column)
    {
        var ps = PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary dict)
        {
            return dict[column];
        }

        var prop = ps.Properties[column];
        if (prop != null)
        {
            return prop.Value;
        }

        var reflectionProp = ps.BaseObject.GetType().GetProperty(column);
        return reflectionProp?.GetValue(ps.BaseObject);
    }
}
