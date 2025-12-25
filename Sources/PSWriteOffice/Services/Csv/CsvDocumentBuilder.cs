using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.CSV;

namespace PSWriteOffice.Services.Csv;

internal static class CsvDocumentBuilder
{
    public static CsvDocument FromObjects(IReadOnlyCollection<object?> items, char delimiter, System.Globalization.CultureInfo? culture, System.Text.Encoding? encoding)
    {
        if (items == null || items.Count == 0)
        {
            throw new ArgumentException("Provide at least one data row.", nameof(items));
        }

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

        var document = new CsvDocument().WithDelimiter(delimiter);
        if (culture != null)
        {
            document.WithCulture(culture);
        }

        if (encoding != null)
        {
            document.WithEncoding(encoding);
        }

        document.WithHeader(columns.ToArray());

        foreach (var item in items)
        {
            if (item == null)
            {
                throw new InvalidOperationException("Data rows cannot contain null entries.");
            }

            var rowValues = new object?[columns.Count];
            for (var i = 0; i < columns.Count; i++)
            {
                rowValues[i] = GetValue(item, columns[i]);
            }

            document.AddRow(rowValues);
        }

        return document;
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

        return ps.BaseObject.GetType()
            .GetProperties()
            .Where(p => p.CanRead)
            .Select(p => p.Name)
            .ToList();
    }

    private static object? GetValue(object item, string column)
    {
        var ps = PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary dict)
        {
            return dict.Contains(column) ? dict[column] : null;
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
