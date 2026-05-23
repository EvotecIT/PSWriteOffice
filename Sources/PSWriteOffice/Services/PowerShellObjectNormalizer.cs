using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Reflection;

namespace PSWriteOffice.Services;

internal static class PowerShellObjectNormalizer
{
    private static readonly ConcurrentDictionary<Type, bool> ClrProjectionCandidateCache = new();

    public static IReadOnlyList<object?> NormalizeItems(IEnumerable<object?> items)
    {
        if (items == null) throw new ArgumentNullException(nameof(items));

        var result = items is ICollection<object?> collection
            ? new List<object?>(collection.Count)
            : new List<object?>();

        foreach (var item in items)
        {
            result.Add(NormalizeItem(item));
        }

        return result;
    }

    public static object? NormalizeItem(object? item)
    {
        if (item == null)
        {
            return null;
        }

        if (item is PSObject psObject)
        {
            return NormalizePSObject(psObject);
        }

        if (item is IDictionary dict)
        {
            return dict;
        }

        if (ClrProjectionCandidateCache.GetOrAdd(item.GetType(), CanUseClrObjectProjection))
        {
            return item;
        }

        return NormalizePSObject(PSObject.AsPSObject(item));
    }

    private static object? NormalizePSObject(PSObject ps)
    {
        if (ps.BaseObject is IDictionary dict)
        {
            return dict;
        }

        Dictionary<string, object?>? result = null;
        foreach (var property in ps.Properties)
        {
            if (property.MemberType != PSMemberTypes.NoteProperty &&
                property.MemberType != PSMemberTypes.Property)
            {
                continue;
            }

            string name = property.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            result ??= new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            result[name] = property.Value;
        }

        if (result != null)
        {
            return result;
        }

        return ps.BaseObject;
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
