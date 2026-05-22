using System;
using System.Collections;
using System.Collections.Generic;
using System.Management.Automation;

namespace PSWriteOffice.Services;

internal static class PowerShellObjectNormalizer
{
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

        var ps = PSObject.AsPSObject(item);
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
}
