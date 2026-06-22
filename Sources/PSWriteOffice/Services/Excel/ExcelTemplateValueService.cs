using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelTemplateValueService
{
    public static Dictionary<string, object?> ConvertValues(Hashtable values)
    {
        var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (DictionaryEntry entry in values)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (!string.IsNullOrWhiteSpace(key))
            {
                result[key!] = entry.Value;
            }
        }

        return result;
    }

    public static IReadOnlyList<IDictionary<string, object?>> ConvertRows(IEnumerable<object?> rows)
    {
        return rows.Select(ConvertRow).ToArray();
    }

    public static ExcelTemplateOptions CreateOptions(string? cultureName, ExcelTemplateMissingValueBehavior? missingValueBehavior, bool throwOnMissing)
    {
        var options = new ExcelTemplateOptions
        {
            ThrowOnMissing = throwOnMissing
        };

        if (missingValueBehavior.HasValue)
        {
            options.MissingValueBehavior = missingValueBehavior.Value;
        }

        if (!string.IsNullOrWhiteSpace(cultureName))
        {
            options.FormatProvider = CultureInfo.GetCultureInfo(cultureName!);
        }

        return options;
    }

    public static string? GetStringValue(IDictionary<string, object?> values, string? propertyName)
    {
        if (string.IsNullOrWhiteSpace(propertyName)
            || !values.TryGetValue(propertyName!, out var value)
            || value == null)
        {
            return null;
        }

        return Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    public static PSObject CreateMarkerRecord(ExcelTemplateMarkerInfo marker, string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("SheetName", marker.SheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", marker.SheetName));
        record.Properties.Add(new PSNoteProperty("Address", marker.CellReference));
        record.Properties.Add(new PSNoteProperty("CellReference", marker.CellReference));
        record.Properties.Add(new PSNoteProperty("Name", marker.Name));
        record.Properties.Add(new PSNoteProperty("Marker", marker.Name));
        record.Properties.Add(new PSNoteProperty("Format", marker.Format));
        record.Properties.Add(new PSNoteProperty("CellText", marker.CellText));
        record.Properties.Add(new PSNoteProperty("IsWholeCell", marker.IsWholeCell));
        record.Properties.Add(new PSNoteProperty("IsBound", marker.IsBound));
        record.Properties.Add(new PSNoteProperty("BoundValueKind", marker.BoundValueKind));
        record.Properties.Add(new PSNoteProperty("BoundValueTypeName", marker.BoundValueTypeName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }

    private static IDictionary<string, object?> ConvertRow(object? row)
    {
        if (row == null)
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        }

        if (row is Hashtable hashtable)
        {
            return ConvertValues(hashtable);
        }

        if (row is IDictionary dictionary)
        {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(key))
                {
                    result[key!] = entry.Value;
                }
            }

            return result;
        }

        var psObject = PSObject.AsPSObject(row);
        var properties = psObject.Properties
            .Where(property => property.IsGettable && property.MemberType is PSMemberTypes.NoteProperty or PSMemberTypes.Property)
            .ToArray();
        var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        foreach (var property in properties)
        {
            values[property.Name] = property.Value;
        }

        return values;
    }
}
