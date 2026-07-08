using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using PSWriteOffice.Services;

namespace PSWriteOffice.Services.Table;

internal static class OfficeTableSpecParser
{
    private static readonly string[] TextKeys = { "Text", "Value", "Content" };
    private static readonly string[] ColumnSpanKeys = { "ColumnSpan", "ColSpan" };
    private static readonly string[] RowSpanKeys = { "RowSpan" };

    public static bool TryCreate(
        IReadOnlyList<object> rows,
        string[]? propertyNames,
        string[]? header,
        out OfficeTableSpec spec)
    {
        spec = null!;
        if (rows.Count == 0)
        {
            return false;
        }

        var hasExplicitRows = rows.All(IsExplicitRow);
        var hasStructuredMarker = rows.Any(ContainsStructuredCellMarker);
        if (!hasExplicitRows && !hasStructuredMarker)
        {
            return false;
        }

        var tableRows = new List<IReadOnlyList<OfficeTableCellSpec>>();
        var hasHeader = false;
        var allowDefaultHeader = header == null;
        if (header is { Length: > 0 })
        {
            tableRows.Add(header.Select(static value => new OfficeTableCellSpec(value)).ToArray());
            hasHeader = true;
        }

        string[]? columns = propertyNames;
        foreach (var row in rows)
        {
            if (TryCreateCell(row, requireExplicitCellShape: true, out var singleCell))
            {
                tableRows.Add(new[] { singleCell });
                continue;
            }

            if (IsExplicitRow(row))
            {
                tableRows.Add(CreateExplicitRow(row));
                continue;
            }

            if (PowerShellObjectNormalizer.TryProjectItem(row, columns, out var projectedColumns, out var values))
            {
                columns ??= projectedColumns;
                if (allowDefaultHeader && !hasHeader && columns.Length > 0)
                {
                    tableRows.Add(columns.Select(static value => new OfficeTableCellSpec(value)).ToArray());
                    hasHeader = true;
                }

                tableRows.Add(values.Select(ToCell).ToArray());
                continue;
            }

            tableRows.Add(new[] { ToCell(row) });
        }

        if (tableRows.Count == 0)
        {
            return false;
        }

        spec = new OfficeTableSpec(tableRows, hasHeader);
        return spec.RowCount > 0 && spec.ColumnCount > 0;
    }

    private static OfficeTableCellSpec[] CreateExplicitRow(object row)
    {
        var cells = new List<OfficeTableCellSpec>();
        foreach (var value in Enumerate(row))
        {
            cells.Add(TryCreateCell(value, requireExplicitCellShape: false, out var spec)
                ? spec
                : ToCell(value));
        }

        return cells.ToArray();
    }

    private static bool ContainsStructuredCellMarker(object row)
    {
        if (TryCreateCell(row, requireExplicitCellShape: true, out var cell) && cell.HasSpan)
        {
            return true;
        }

        if (!IsExplicitRow(row))
        {
            return false;
        }

        foreach (var value in Enumerate(row))
        {
            if (TryCreateCell(value, requireExplicitCellShape: false, out cell) && cell.HasSpan)
            {
                return true;
            }
        }

        return false;
    }

    private static OfficeTableCellSpec ToCell(object? value)
        => new(Convert.ToString(UnwrapPSObject(value), CultureInfo.InvariantCulture));

    private static bool TryCreateCell(object? value, bool requireExplicitCellShape, out OfficeTableCellSpec spec)
    {
        value = UnwrapPSObject(value);
        if (value is OfficeTableCellSpec typed)
        {
            spec = typed;
            return true;
        }

        var hasColumnSpan = TryGetValue(value, ColumnSpanKeys, out var columnSpanValue);
        var hasRowSpan = TryGetValue(value, RowSpanKeys, out var rowSpanValue);
        var hasSpan = hasColumnSpan || hasRowSpan;
        var hasText = TryGetValue(value, TextKeys, out var textValue);
        if (requireExplicitCellShape && !hasSpan)
        {
            spec = null!;
            return false;
        }

        if (!hasSpan && !hasText)
        {
            spec = null!;
            return false;
        }

        var text = hasText
            ? Convert.ToString(UnwrapPSObject(textValue), CultureInfo.InvariantCulture)
            : string.Empty;
        var columnSpan = hasColumnSpan
            ? ConvertToPositiveInt(columnSpanValue, "ColumnSpan")
            : 1;
        var rowSpan = hasRowSpan
            ? ConvertToPositiveInt(rowSpanValue, "RowSpan")
            : 1;

        spec = new OfficeTableCellSpec(text, columnSpan, rowSpan);
        return true;
    }

    private static bool IsExplicitRow(object row)
    {
        row = UnwrapPSObject(row)!;
        return row is IEnumerable and not string and not IDictionary and not DataTable and not DataView and not IDataReader and not DataSet &&
               !IsGenericDictionary(row);
    }

    private static IEnumerable<object?> Enumerate(object row)
    {
        foreach (var value in (IEnumerable)UnwrapPSObject(row)!)
        {
            yield return value;
        }
    }

    private static bool TryGetValue(
        object? source,
        IEnumerable<string> keys,
        out object? value)
    {
        if (source is IDictionary dictionary)
        {
            foreach (var key in keys)
            {
                if (dictionary.Contains(key))
                {
                    value = dictionary[key];
                    return true;
                }

                foreach (DictionaryEntry entry in dictionary)
                {
                    if (string.Equals(Convert.ToString(entry.Key, CultureInfo.InvariantCulture), key, StringComparison.OrdinalIgnoreCase))
                    {
                        value = entry.Value;
                        return true;
                    }
                }
            }

            value = null;
            return false;
        }

        if (source is null)
        {
            value = null;
            return false;
        }

        var psObject = PSObject.AsPSObject(source);
        if (psObject.BaseObject is IDictionary baseDictionary)
        {
            return TryGetValue(baseDictionary, keys, out value);
        }

        foreach (var key in keys)
        {
            var property = psObject.Properties[key];
            if (property == null ||
                !property.IsGettable ||
                property.MemberType is not (PSMemberTypes.NoteProperty or PSMemberTypes.Property or PSMemberTypes.AliasProperty))
            {
                continue;
            }

            try
            {
                value = property.Value;
                return true;
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                value = null;
                return false;
            }
        }

        value = null;
        return false;
    }

    private static int ConvertToPositiveInt(object? value, string name)
    {
        value = UnwrapPSObject(value);
        var result = (int)LanguagePrimitives.ConvertTo(value, typeof(int), CultureInfo.InvariantCulture);
        if (result < 1)
        {
            throw new ArgumentOutOfRangeException(name, "Span values must be at least 1.");
        }

        return result;
    }

    private static object? UnwrapPSObject(object? value)
        => value is PSObject psObject ? psObject.BaseObject : value;

    private static bool IsGenericDictionary(object value)
    {
        foreach (var interfaceType in value.GetType().GetInterfaces())
        {
            if (!interfaceType.IsGenericType)
            {
                continue;
            }

            var genericDefinition = interfaceType.GetGenericTypeDefinition();
            if (genericDefinition == typeof(IDictionary<,>) ||
                genericDefinition == typeof(IReadOnlyDictionary<,>))
            {
                return true;
            }
        }

        return false;
    }
}
