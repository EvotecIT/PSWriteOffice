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

        if (ClrProjectionCandidateCache.GetOrAdd(item.GetType(), CanUseClrObjectProjection))
        {
            return item;
        }

        return NormalizePSObject(PSObject.AsPSObject(item), options);
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
