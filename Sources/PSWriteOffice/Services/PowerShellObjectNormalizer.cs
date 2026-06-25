using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
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

        if (item is PSObject psObject)
        {
            return NormalizePSObject(psObject, options);
        }

        if (item is IDictionary dict)
        {
            return dict;
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

    private static bool ShouldExportProperty(PSPropertyInfo property)
    {
        if (!property.IsGettable)
        {
            return false;
        }

        return property.MemberType == PSMemberTypes.NoteProperty ||
            property.MemberType == PSMemberTypes.Property ||
            property.MemberType == PSMemberTypes.ScriptProperty ||
            property.MemberType == PSMemberTypes.CodeProperty;
    }

    private static object? NormalizeCellValue(object? value, PowerShellObjectNormalizerOptions options)
    {
        if (value == null || value is string || !options.NormalizeCollectionValues)
        {
            return value;
        }

        if (value is bool ||
            value is char ||
            value is byte ||
            value is sbyte ||
            value is short ||
            value is ushort ||
            value is int ||
            value is uint ||
            value is long ||
            value is ulong ||
            value is float ||
            value is double ||
            value is decimal ||
            value is DateTime ||
            value is DateTimeOffset ||
            value is TimeSpan ||
            value is Guid)
        {
            return value;
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
        if (type == typeof(string) ||
            type.IsPrimitive ||
            type.IsEnum ||
            type == typeof(decimal) ||
            type == typeof(DateTime) ||
            type == typeof(DateTimeOffset) ||
            type == typeof(TimeSpan) ||
            typeof(IEnumerable).IsAssignableFrom(type))
        {
            return false;
        }

        return type.GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Any(static property => property.CanRead && property.GetIndexParameters().Length == 0);
    }

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
}
