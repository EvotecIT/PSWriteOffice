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
    private static readonly string[] ColumnSpanKeys = { "ColumnSpan", "ColSpan", "Columns", "Span" };
    private static readonly string[] RowSpanKeys = { "RowSpan", "Rows" };

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
        if (header is { Length: > 0 })
        {
            tableRows.Add(header.Select(static value => new OfficeTableCellSpec(value)).ToArray());
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
                tableRows.Add(values.Select(ToCell).ToArray());
                continue;
            }

            tableRows.Add(new[] { ToCell(row) });
        }

        if (tableRows.Count == 0)
        {
            return false;
        }

        spec = new OfficeTableSpec(tableRows);
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

        if (!TryGetPropertyBag(value, out var properties))
        {
            spec = null!;
            return false;
        }

        var hasSpan = TryGetValue(properties, ColumnSpanKeys, out _) ||
            TryGetValue(properties, RowSpanKeys, out _);
        var hasText = TryGetValue(properties, TextKeys, out _);
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

        var text = TryGetValue(properties, TextKeys, out var textValue)
            ? Convert.ToString(UnwrapPSObject(textValue), CultureInfo.InvariantCulture)
            : string.Empty;
        var columnSpan = TryGetValue(properties, ColumnSpanKeys, out var columnSpanValue)
            ? ConvertToPositiveInt(columnSpanValue, "ColumnSpan")
            : 1;
        var rowSpan = TryGetValue(properties, RowSpanKeys, out var rowSpanValue)
            ? ConvertToPositiveInt(rowSpanValue, "RowSpan")
            : 1;

        spec = new OfficeTableCellSpec(text, columnSpan, rowSpan);
        return true;
    }

    private static bool IsExplicitRow(object row)
        => row is IEnumerable and not string and not IDictionary and not DataTable and not DataView and not IDataReader and not DataSet &&
           !IsGenericDictionary(row);

    private static IEnumerable<object?> Enumerate(object row)
    {
        foreach (var value in (IEnumerable)row)
        {
            yield return value;
        }
    }

    private static bool TryGetPropertyBag(object? value, out Dictionary<string, object?> properties)
    {
        properties = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    properties[key!] = entry.Value;
                }
            }

            return properties.Count > 0;
        }

        if (value is null)
        {
            return false;
        }

        var psObject = PSObject.AsPSObject(value);
        if (psObject.BaseObject is IDictionary baseDictionary)
        {
            return TryGetPropertyBag(baseDictionary, out properties);
        }

        foreach (var property in psObject.Properties)
        {
            if (!property.IsGettable ||
                property.MemberType is not (PSMemberTypes.NoteProperty or PSMemberTypes.Property or PSMemberTypes.AliasProperty))
            {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(property.Name))
            {
                properties[property.Name] = property.Value;
            }
        }

        return properties.Count > 0;
    }

    private static bool TryGetValue(
        IReadOnlyDictionary<string, object?> properties,
        IEnumerable<string> keys,
        out object? value)
    {
        foreach (var key in keys)
        {
            if (properties.TryGetValue(key, out value))
            {
                return true;
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
