using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Services;

internal static class PowerShellObjectNormalizer
{
    private static readonly ConcurrentDictionary<Type, bool> ClrProjectionCandidateCache = new();
    private static readonly ConcurrentDictionary<Type, ClrProjectionPlan> ClrProjectionPlanCache = new();

    public static IReadOnlyList<object?> NormalizeItems(IEnumerable<object?> items, PowerShellObjectNormalizerOptions? options = null)
    {
        if (items == null) throw new ArgumentNullException(nameof(items));

        options ??= PowerShellObjectNormalizerOptions.Default;
        var result = items is ICollection<object?> collection
            ? new List<object?>(collection.Count)
            : new List<object?>();

        foreach (var item in items)
        {
            result.Add(NormalizeItem(item, options));
        }

        return result;
    }

    public static object? NormalizeItem(object? item, PowerShellObjectNormalizerOptions? options = null)
    {
        options ??= PowerShellObjectNormalizerOptions.Default;
        if (item == null)
        {
            return null;
        }

        if (TryGetDataRow(item, out var dataRow))
        {
            return ProjectDataRowDictionary(dataRow, options);
        }

        if (item is PSObject psObject)
        {
            return NormalizePSObject(psObject, options);
        }

        if (item is IDictionary dict)
        {
            return dict;
        }

        if (IsScalarClrType(item.GetType()))
        {
            return NormalizeCellValue(item, options);
        }

        if (TryGetClrProjectionPlan(item.GetType(), out _))
        {
            return item;
        }

        return NormalizePSObject(PSObject.AsPSObject(item), options);
    }

    public static bool TryProjectItem(object? item, string[]? columns, out string[] projectedColumns, out object?[] values, PowerShellObjectNormalizerOptions? options = null)
    {
        options ??= PowerShellObjectNormalizerOptions.Default;
        projectedColumns = columns ?? Array.Empty<string>();
        values = Array.Empty<object?>();

        if (item == null)
        {
            return false;
        }

        if (TryProjectScalar(item, columns, out projectedColumns, out values, options))
        {
            return true;
        }

        if (TryGetDataRow(item, out var dataRow))
        {
            ProjectDataRow(dataRow, columns, out projectedColumns, out values, options);
            return true;
        }

        if (item is PSObject psObject)
        {
            return TryProjectPSObject(psObject, columns, out projectedColumns, out values, options);
        }

        if (item is IDictionary dictionary)
        {
            ProjectDictionary(dictionary, columns, out projectedColumns, out values);
            return true;
        }

        if (TryGetClrProjectionPlan(item.GetType(), out var plan))
        {
            ProjectClrObject(item, plan, columns, out projectedColumns, out values, options);
            return true;
        }

        return TryProjectPSObject(PSObject.AsPSObject(item), columns, out projectedColumns, out values, options);
    }

    public static bool TryProjectItemInto(object? item, string[] columns, object?[] values, PowerShellObjectNormalizerOptions? options = null)
    {
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (values.Length != columns.Length)
        {
            throw new ArgumentException("The value buffer length must match the column count.", nameof(values));
        }

        options ??= PowerShellObjectNormalizerOptions.Default;

        if (item == null)
        {
            return false;
        }

        if (TryProjectScalarInto(item, columns, values, options))
        {
            return true;
        }

        if (TryGetDataRow(item, out var dataRow))
        {
            ProjectDataRowInto(dataRow, columns, values, options);
            return true;
        }

        if (item is PSObject psObject)
        {
            ProjectPSObjectInto(psObject, columns, values, options);
            return true;
        }

        if (item is IDictionary dictionary)
        {
            ProjectDictionaryInto(dictionary, columns, values);
            return true;
        }

        if (TryGetClrProjectionPlan(item.GetType(), out var plan))
        {
            ProjectClrObjectInto(item, plan, columns, values, options);
            return true;
        }

        ProjectPSObjectInto(PSObject.AsPSObject(item), columns, values, options);
        return true;
    }

    public static bool TryProjectPSObjectIntoKnownColumns(object? item, string[] columns, object?[] values, PowerShellObjectNormalizerOptions? options = null)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(values);
#else
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (values == null) throw new ArgumentNullException(nameof(values));
#endif

        if (values.Length != columns.Length)
        {
            throw new ArgumentException("The value buffer length must match the column count.", nameof(values));
        }

        options ??= PowerShellObjectNormalizerOptions.Default;

        if (item == null || item is IDictionary)
        {
            return false;
        }

        if (item is not PSObject && item.GetType().FullName != "System.Management.Automation.PSCustomObject")
        {
            return false;
        }

        var ps = item as PSObject ?? PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary)
        {
            return false;
        }

        if (ps.BaseObject != null &&
            IsScalarClrType(ps.BaseObject.GetType()) &&
            !HasExportableExtendedProperties(ps))
        {
            return false;
        }

        if (TryProjectAlignedPSObject(ps, columns, values, options))
        {
            return true;
        }

        for (var i = 0; i < columns.Length; i++)
        {
            var property = ps.Properties[columns[i]];
            if (property == null || !ShouldExportProperty(property))
            {
                values[i] = null;
                continue;
            }

            try
            {
                values[i] = NormalizeCellValue(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(property.Name, exception);
                values[i] = options.IncludeUnexportableProperties
                    ? options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception) ?? string.Empty
                    : null;
            }
        }

        return true;
    }

    public static bool TryProjectPSObjectTextIntoKnownColumns(object? item, string[] columns, string?[] values, PowerShellObjectNormalizerOptions? options = null)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(columns);
        ArgumentNullException.ThrowIfNull(values);
#else
        if (columns == null) throw new ArgumentNullException(nameof(columns));
        if (values == null) throw new ArgumentNullException(nameof(values));
#endif

        if (values.Length != columns.Length)
        {
            throw new ArgumentException("The value buffer length must match the column count.", nameof(values));
        }

        options ??= PowerShellObjectNormalizerOptions.Default;

        if (item == null || item is IDictionary)
        {
            return false;
        }

        if (item is not PSObject && item.GetType().FullName != "System.Management.Automation.PSCustomObject")
        {
            return false;
        }

        var ps = item as PSObject ?? PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary)
        {
            return false;
        }

        if (ps.BaseObject != null &&
            IsScalarClrType(ps.BaseObject.GetType()) &&
            !HasExportableExtendedProperties(ps))
        {
            return false;
        }

        for (var i = 0; i < columns.Length; i++)
        {
            var property = ps.Properties[columns[i]];
            if (property == null || !ShouldExportProperty(property))
            {
                values[i] = null;
                continue;
            }

            try
            {
                values[i] = NormalizeCellValueToText(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(property.Name, exception);
                values[i] = options.IncludeUnexportableProperties
                    ? options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception)?.ToString() ?? string.Empty
                    : null;
            }
        }

        return true;
    }

    public static bool TryPreparePSObjectTextProjection(object? item, out PSObject? ps)
    {
        ps = null;
        if (item == null || item is IDictionary)
        {
            return false;
        }

        if (item is not PSObject && item.GetType().FullName != "System.Management.Automation.PSCustomObject")
        {
            return false;
        }

        ps = item as PSObject ?? PSObject.AsPSObject(item);
        if (ps.BaseObject is IDictionary)
        {
            ps = null;
            return false;
        }

        if (ps.BaseObject != null &&
            IsScalarClrType(ps.BaseObject.GetType()) &&
            !HasExportableExtendedProperties(ps))
        {
            ps = null;
            return false;
        }

        return true;
    }

    public static string? ProjectPSObjectTextValue(PSObject ps, string column, PowerShellObjectNormalizerOptions options)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(ps);
        ArgumentNullException.ThrowIfNull(column);
#else
        if (ps == null) throw new ArgumentNullException(nameof(ps));
        if (column == null) throw new ArgumentNullException(nameof(column));
#endif

        var property = ps.Properties[column];
        if (property == null || !ShouldExportProperty(property))
        {
            return null;
        }

        return ProjectPSObjectPropertyTextValue(property, options);
    }

    public static object? ProjectPSObjectValue(PSObject ps, string column, PowerShellObjectNormalizerOptions options)
    {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(ps);
        ArgumentNullException.ThrowIfNull(column);
#else
        if (ps == null) throw new ArgumentNullException(nameof(ps));
        if (column == null) throw new ArgumentNullException(nameof(column));
#endif

        var property = ps.Properties[column];
        if (property == null || !ShouldExportProperty(property))
        {
            return null;
        }

        return ProjectPSObjectPropertyValue(property, options);
    }

    private static bool TryProjectAlignedPSObject(PSObject ps, string[] columns, object?[] values, PowerShellObjectNormalizerOptions options)
    {
        var index = 0;
        foreach (var property in ps.Properties)
        {
            if (!ShouldExportProperty(property))
            {
                continue;
            }

            if (!IsAlignedColumn(property.Name, columns, index))
            {
                return false;
            }

            if (property.MemberType == PSMemberTypes.NoteProperty &&
                TryGetFastCellValue(property.Value, options, out values[index]))
            {
                index++;
                continue;
            }

            try
            {
                values[index] = NormalizeCellValue(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(property.Name, exception);
                values[index] = options.IncludeUnexportableProperties
                    ? options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception) ?? string.Empty
                    : null;
            }

            index++;
        }

        return index == columns.Length;
    }

    private static bool IsAlignedColumn(string propertyName, string[] columns, int index)
    {
        return index < columns.Length &&
            (string.Equals(propertyName, columns[index], StringComparison.Ordinal) ||
             string.Equals(propertyName, columns[index], StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryProjectScalar(object item, string[]? columns, out string[] projectedColumns, out object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (item is PSObject psObject && HasExportableExtendedProperties(psObject))
        {
            projectedColumns = columns ?? Array.Empty<string>();
            values = Array.Empty<object?>();
            return false;
        }

        var value = item is PSObject { BaseObject: { } baseObject } && IsScalarClrType(baseObject.GetType())
            ? baseObject
            : item;

        if (!IsScalarClrType(value.GetType()))
        {
            projectedColumns = columns ?? Array.Empty<string>();
            values = Array.Empty<object?>();
            return false;
        }

        projectedColumns = columns ?? new[] { "Value" };
        values = new object?[projectedColumns.Length];
        var targetIndex = ResolveScalarProjectionIndex(projectedColumns);
        if (targetIndex < 0)
        {
            return false;
        }

        if (values.Length > 0)
        {
            values[targetIndex] = NormalizeCellValue(value, options);
        }

        return true;
    }

    private static bool TryProjectScalarInto(object item, string[] columns, object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (item is PSObject psObject && HasExportableExtendedProperties(psObject))
        {
            return false;
        }

        var value = item is PSObject { BaseObject: { } baseObject } && IsScalarClrType(baseObject.GetType())
            ? baseObject
            : item;

        if (!IsScalarClrType(value.GetType()))
        {
            return false;
        }

        var targetIndex = ResolveScalarProjectionIndex(columns);
        if (targetIndex < 0)
        {
            return false;
        }

        Array.Clear(values, 0, values.Length);
        values[targetIndex] = NormalizeCellValue(value, options);
        return true;
    }

    private static int ResolveScalarProjectionIndex(IReadOnlyList<string> columns)
    {
        if (columns.Count == 0)
        {
            return -1;
        }

        if (columns.Count == 1)
        {
            return 0;
        }

        for (var i = 0; i < columns.Count; i++)
        {
            if (string.Equals(columns[i], "Value", StringComparison.OrdinalIgnoreCase))
            {
                return i;
            }
        }

        return -1;
    }

    private static bool HasExportableExtendedProperties(PSObject psObject)
        => psObject.Properties.Any(static property =>
            IsExtendedProperty(property) &&
            !string.IsNullOrWhiteSpace(property.Name));

    private static bool IsExtendedProperty(PSPropertyInfo property)
        => property.IsInstance &&
           property.IsGettable &&
           (property.MemberType == PSMemberTypes.AliasProperty ||
            property.MemberType == PSMemberTypes.NoteProperty ||
            property.MemberType == PSMemberTypes.ScriptProperty ||
            property.MemberType == PSMemberTypes.CodeProperty);

    private static object? NormalizePSObject(PSObject ps, PowerShellObjectNormalizerOptions options)
    {
        if (ps.BaseObject is IDictionary dict)
        {
            return dict;
        }

        Dictionary<string, object?>? result = null;
        foreach (var property in ps.Properties)
        {
            if (!ShouldExportProperty(property))
            {
                continue;
            }

            string name = property.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            result ??= new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            try
            {
                result[name] = NormalizeCellValue(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{name}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(name, exception);
                if (options.IncludeUnexportableProperties)
                {
                    result[name] = options.UnexportablePropertyValueFactory?.Invoke(name, exception) ?? string.Empty;
                }
            }
        }

        if (result != null)
        {
            return result;
        }

        return ps.BaseObject;
    }

    private static bool TryProjectPSObject(PSObject ps, string[]? columns, out string[] projectedColumns, out object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (ps.BaseObject is IDictionary dictionary)
        {
            ProjectDictionary(dictionary, columns, out projectedColumns, out values);
            return true;
        }

        if (columns == null)
        {
            var names = new List<string>();
            var rowValues = new List<object?>();
            foreach (var property in ps.Properties)
            {
                if (!ShouldExportProperty(property) || string.IsNullOrWhiteSpace(property.Name))
                {
                    continue;
                }

                try
                {
                    var value = NormalizeCellValue(property.Value, options);
                    names.Add(property.Name);
                    rowValues.Add(value);
                }
                catch (Exception exception) when (exception is not PipelineStoppedException)
                {
                    if (options.PropertyErrorAction == ActionPreference.Stop)
                    {
                        throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
                    }

                    options.PropertyErrorCallback?.Invoke(property.Name, exception);
                    if (options.IncludeUnexportableProperties)
                    {
                        names.Add(property.Name);
                        rowValues.Add(options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception) ?? string.Empty);
                    }
                }
            }

            projectedColumns = names.ToArray();
            values = rowValues.ToArray();
            return names.Count > 0;
        }

        projectedColumns = columns;
        values = new object?[columns.Length];
        for (var i = 0; i < columns.Length; i++)
        {
            var property = ps.Properties[columns[i]];
            if (property == null || !ShouldExportProperty(property))
            {
                values[i] = null;
                continue;
            }

            try
            {
                values[i] = NormalizeCellValue(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{columns[i]}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(columns[i], exception);
                if (options.IncludeUnexportableProperties)
                {
                    values[i] = options.UnexportablePropertyValueFactory?.Invoke(columns[i], exception) ?? string.Empty;
                }
            }
        }

        return true;
    }

    private static void ProjectPSObjectInto(PSObject ps, string[] columns, object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (ps.BaseObject is IDictionary dictionary)
        {
            ProjectDictionaryInto(dictionary, columns, values);
            return;
        }

        for (var i = 0; i < columns.Length; i++)
        {
            var property = ps.Properties[columns[i]];
            if (property == null || !ShouldExportProperty(property))
            {
                values[i] = null;
                continue;
            }

            try
            {
                values[i] = NormalizeCellValue(property.Value, options);
            }
            catch (Exception exception) when (exception is not PipelineStoppedException)
            {
                if (options.PropertyErrorAction == ActionPreference.Stop)
                {
                    throw new InvalidOperationException($"Unable to read PowerShell property '{columns[i]}'.", exception);
                }

                options.PropertyErrorCallback?.Invoke(columns[i], exception);
                if (options.IncludeUnexportableProperties)
                {
                    values[i] = options.UnexportablePropertyValueFactory?.Invoke(columns[i], exception) ?? string.Empty;
                }
                else
                {
                    values[i] = null;
                }
            }
        }
    }

    private static void ProjectDictionary(IDictionary dictionary, string[]? columns, out string[] projectedColumns, out object?[] values)
    {
        if (columns == null)
        {
            var names = new List<string>();
            var rowValues = new List<object?>();
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = entry.Key?.ToString();
                if (string.IsNullOrWhiteSpace(key))
                {
                    continue;
                }

                names.Add(key!);
                rowValues.Add(entry.Value);
            }

            projectedColumns = names.ToArray();
            values = rowValues.ToArray();
            return;
        }

        projectedColumns = columns;
        values = new object?[columns.Length];
        for (var i = 0; i < columns.Length; i++)
        {
            values[i] = GetDictionaryValue(dictionary, columns[i]);
        }
    }

    private static void ProjectDictionaryInto(IDictionary dictionary, string[] columns, object?[] values)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            values[i] = GetDictionaryValue(dictionary, columns[i]);
        }
    }

    private static Dictionary<string, object?> ProjectDataRowDictionary(DataRow row, PowerShellObjectNormalizerOptions options)
    {
        var columns = row.Table.Columns;
        var result = new Dictionary<string, object?>(columns.Count, StringComparer.OrdinalIgnoreCase);
        foreach (DataColumn column in columns)
        {
            result[column.ColumnName] = NormalizeCellValue(row[column], options);
        }

        return result;
    }

    private static bool TryGetDataRow(object item, out DataRow row)
    {
        if (item is PSObject psObject)
        {
            item = psObject.BaseObject;
        }

        if (item is DataRow dataRow)
        {
            row = dataRow;
            return true;
        }

        if (item is DataRowView dataRowView)
        {
            row = dataRowView.Row;
            return true;
        }

        row = null!;
        return false;
    }

    private static void ProjectDataRow(DataRow row, string[]? columns, out string[] projectedColumns, out object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (columns == null)
        {
            var tableColumns = row.Table.Columns;
            projectedColumns = new string[tableColumns.Count];
            values = new object?[tableColumns.Count];
            for (var i = 0; i < tableColumns.Count; i++)
            {
                var column = tableColumns[i];
                projectedColumns[i] = column.ColumnName;
                values[i] = NormalizeCellValue(row[column], options);
            }

            return;
        }

        projectedColumns = columns;
        values = new object?[columns.Length];
        ProjectDataRowInto(row, columns, values, options);
    }

    private static void ProjectDataRowInto(DataRow row, string[] columns, object?[] values, PowerShellObjectNormalizerOptions options)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            var column = GetDataColumn(row.Table.Columns, columns[i]);
            values[i] = column == null ? null : NormalizeCellValue(row[column], options);
        }
    }

    private static void ProjectClrObject(object item, ClrProjectionPlan plan, string[]? columns, out string[] projectedColumns, out object?[] values, PowerShellObjectNormalizerOptions options)
    {
        if (columns == null)
        {
            var names = new List<string>(plan.Properties.Count);
            var rowValues = new List<object?>(plan.Properties.Count);
            foreach (var property in plan.Properties)
            {
                if (TryReadClrProperty(item, property, property.Name, options, out var value))
                {
                    names.Add(property.Name);
                    rowValues.Add(value);
                }
            }

            projectedColumns = names.ToArray();
            values = rowValues.ToArray();
            return;
        }

        projectedColumns = columns;
        values = new object?[columns.Length];
        for (var i = 0; i < columns.Length; i++)
        {
            var property = plan.PropertiesByName.TryGetValue(columns[i], out var candidate)
                ? candidate
                : null;
            if (property == null || !TryReadClrProperty(item, property, columns[i], options, out var value))
            {
                values[i] = null;
                continue;
            }

            values[i] = value;
        }
    }

    private static void ProjectClrObjectInto(object item, ClrProjectionPlan plan, string[] columns, object?[] values, PowerShellObjectNormalizerOptions options)
    {
        for (var i = 0; i < columns.Length; i++)
        {
            var property = plan.PropertiesByName.TryGetValue(columns[i], out var candidate)
                ? candidate
                : null;
            if (property == null || !TryReadClrProperty(item, property, columns[i], options, out var value))
            {
                values[i] = null;
                continue;
            }

            values[i] = value;
        }
    }

    private static bool TryReadClrProperty(object item, PropertyInfo property, string name, PowerShellObjectNormalizerOptions options, out object? value)
    {
        try
        {
            value = NormalizeCellValue(property.GetValue(item), options);
            return true;
        }
        catch (Exception exception) when (exception is not PipelineStoppedException)
        {
            exception = UnwrapClrPropertyException(exception);
            if (options.PropertyErrorAction == ActionPreference.Stop)
            {
                throw new InvalidOperationException($"Unable to read CLR property '{name}'.", exception);
            }

            options.PropertyErrorCallback?.Invoke(name, exception);
            if (options.IncludeUnexportableProperties)
            {
                value = options.UnexportablePropertyValueFactory?.Invoke(name, exception) ?? string.Empty;
                return true;
            }

            value = null;
            return false;
        }
    }

    private static Exception UnwrapClrPropertyException(Exception exception) =>
        exception is TargetInvocationException { InnerException: not null } ? exception.InnerException : exception;

    private static object? GetDictionaryValue(IDictionary dictionary, string column)
    {
        if (dictionary.Contains(column))
        {
            return dictionary[column];
        }

        foreach (DictionaryEntry entry in dictionary)
        {
            if (string.Equals(entry.Key?.ToString(), column, StringComparison.OrdinalIgnoreCase))
            {
                return entry.Value;
            }
        }

        return null;
    }

    private static DataColumn? GetDataColumn(DataColumnCollection columns, string columnName)
    {
        if (columns.Contains(columnName))
        {
            return columns[columnName];
        }

        foreach (DataColumn column in columns)
        {
            if (string.Equals(column.ColumnName, columnName, StringComparison.OrdinalIgnoreCase))
            {
                return column;
            }
        }

        return null;
    }

    private static bool ShouldExportProperty(PSPropertyInfo property)
    {
        if (!property.IsGettable)
        {
            return false;
        }

        return property.MemberType == PSMemberTypes.AliasProperty ||
            property.MemberType == PSMemberTypes.NoteProperty ||
            property.MemberType == PSMemberTypes.Property ||
            property.MemberType == PSMemberTypes.ScriptProperty ||
            property.MemberType == PSMemberTypes.CodeProperty;
    }

    private static object? NormalizeCellValue(object? value, PowerShellObjectNormalizerOptions options)
    {
        if (TryGetFastCellValue(value, options, out var normalized))
        {
            return normalized;
        }

        if (value is IDictionary)
        {
            return value;
        }

        if (value is IEnumerable enumerable)
        {
            var values = new List<string>();
            foreach (var item in enumerable)
            {
                values.Add(ConvertValueToString(item));
            }

            return string.Join(options.CollectionSeparator, values);
        }

        return value;
    }

    private static string? NormalizeCellValueToText(object? value, PowerShellObjectNormalizerOptions options)
    {
        if (value == null)
        {
            return null;
        }

        if (value is string text)
        {
            return text;
        }

        if (value is bool boolValue)
        {
            return boolValue ? "True" : "False";
        }

        if (value is decimal or int or DateTime or double or long or DateTimeOffset or Guid or TimeSpan or float or byte or sbyte or short or ushort or uint or ulong)
        {
            return value is IFormattable formattable
                ? formattable.ToString(null, options.Culture)
                : value.ToString();
        }

        if (value is char charValue)
        {
            return charValue.ToString();
        }

        if (value is IDictionary)
        {
            return value.ToString();
        }

        if (options.NormalizeCollectionValues && value is IEnumerable enumerable)
        {
            var values = new List<string>();
            foreach (var item in enumerable)
            {
                values.Add(ConvertValueToString(item));
            }

            return string.Join(options.CollectionSeparator, values);
        }

        return value is IFormattable fallbackFormattable
            ? fallbackFormattable.ToString(null, options.Culture)
            : value.ToString();
    }

    private static string? ProjectPSObjectPropertyTextValue(PSPropertyInfo property, PowerShellObjectNormalizerOptions options)
    {
        try
        {
            return NormalizeCellValueToText(property.Value, options);
        }
        catch (Exception exception) when (exception is not PipelineStoppedException)
        {
            if (options.PropertyErrorAction == ActionPreference.Stop)
            {
                throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
            }

            options.PropertyErrorCallback?.Invoke(property.Name, exception);
            return options.IncludeUnexportableProperties
                ? options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception)?.ToString() ?? string.Empty
                : null;
        }
    }

    private static object? ProjectPSObjectPropertyValue(PSPropertyInfo property, PowerShellObjectNormalizerOptions options)
    {
        try
        {
            return NormalizeCellValue(property.Value, options);
        }
        catch (Exception exception) when (exception is not PipelineStoppedException)
        {
            if (options.PropertyErrorAction == ActionPreference.Stop)
            {
                throw new InvalidOperationException($"Unable to read PowerShell property '{property.Name}'.", exception);
            }

            options.PropertyErrorCallback?.Invoke(property.Name, exception);
            return options.IncludeUnexportableProperties
                ? options.UnexportablePropertyValueFactory?.Invoke(property.Name, exception) ?? string.Empty
                : null;
        }
    }

    private static bool TryGetFastCellValue(object? value, PowerShellObjectNormalizerOptions options, out object? normalized)
    {
        normalized = value;
        if (value == null || value is string || !options.NormalizeCollectionValues)
        {
            return true;
        }

        if (value is bool boolValue)
        {
            normalized = options.FormatScalarValuesAsText ? boolValue ? "True" : "False" : value;
            return true;
        }

        if (value is decimal or int or DateTime or double or long or DateTimeOffset or Guid or TimeSpan or float or byte or sbyte or short or ushort or uint or ulong)
        {
            if (options.FormatScalarValuesAsText)
            {
                normalized = value is IFormattable formattable
                    ? formattable.ToString(null, options.Culture)
                    : value.ToString();
            }

            return true;
        }

        if (value is char charValue)
        {
            normalized = options.FormatScalarValuesAsText ? charValue.ToString() : value;
            return true;
        }

        return false;
    }

    private static string ConvertValueToString(object? value)
    {
        if (value == null)
        {
            return string.Empty;
        }

        return LanguagePrimitives.ConvertTo(value, typeof(string), CultureInfo.InvariantCulture) as string ?? string.Empty;
    }

    private static bool TryGetClrProjectionPlan(Type type, out ClrProjectionPlan plan)
    {
        if (!ClrProjectionCandidateCache.GetOrAdd(type, CanUseClrObjectProjection))
        {
            plan = ClrProjectionPlan.Empty;
            return false;
        }

        plan = ClrProjectionPlanCache.GetOrAdd(type, CreateClrProjectionPlan);
        return plan.Properties.Count > 0;
    }

    private static bool CanUseClrObjectProjection(Type type)
    {
        if (IsScalarClrType(type) ||
            typeof(DataRow).IsAssignableFrom(type) ||
            typeof(DataRowView).IsAssignableFrom(type) ||
            typeof(IEnumerable).IsAssignableFrom(type))
        {
            return false;
        }

        return type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Any(static property => property.CanRead && property.GetIndexParameters().Length == 0);
    }

    private static bool IsScalarClrType(Type type) =>
        type == typeof(string) ||
        type.IsPrimitive ||
        type.IsEnum ||
        type == typeof(decimal) ||
        type == typeof(DateTime) ||
        type == typeof(DateTimeOffset) ||
        type == typeof(Guid) ||
        type == typeof(TimeSpan);

    private static ClrProjectionPlan CreateClrProjectionPlan(Type type)
    {
        var properties = type
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(static property => property.CanRead && property.GetIndexParameters().Length == 0)
            .OrderBy(static property => property.MetadataToken)
            .ToArray();

        return new ClrProjectionPlan(properties);
    }

    private sealed class ClrProjectionPlan
    {
        public static readonly ClrProjectionPlan Empty = new(Array.Empty<PropertyInfo>());

        public ClrProjectionPlan(IReadOnlyList<PropertyInfo> properties)
        {
            Properties = properties;
            var columnNames = new string[properties.Count];
            var propertiesByName = new Dictionary<string, PropertyInfo>(properties.Count, StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < properties.Count; i++)
            {
                var property = properties[i];
                columnNames[i] = property.Name;
                propertiesByName[property.Name] = property;
            }

            ColumnNames = columnNames;
            PropertiesByName = propertiesByName;
        }

        public IReadOnlyList<PropertyInfo> Properties { get; }

        public string[] ColumnNames { get; }

        public IReadOnlyDictionary<string, PropertyInfo> PropertiesByName { get; }
    }

}

internal sealed class PowerShellObjectNormalizerOptions
{
    internal static readonly PowerShellObjectNormalizerOptions Default = new();

    public bool IncludeUnexportableProperties { get; set; }

    public ActionPreference PropertyErrorAction { get; set; } = ActionPreference.SilentlyContinue;

    public Action<string, Exception>? PropertyErrorCallback { get; set; }

    public Func<string, Exception, object?>? UnexportablePropertyValueFactory { get; set; }

    public bool NormalizeCollectionValues { get; set; } = true;

    public string CollectionSeparator { get; set; } = ", ";

    public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;

    public bool FormatScalarValuesAsText { get; set; }
}
