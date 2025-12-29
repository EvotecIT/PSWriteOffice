using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;

namespace PSWriteOffice.Services;

internal static class PowerShellObjectNormalizer
{
    public static IReadOnlyList<object?> NormalizeItems(IEnumerable<object?> items)
    {
        if (items == null) throw new ArgumentNullException(nameof(items));
        return items.Select(NormalizeItem).ToList();
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

        var properties = ps.Properties
            .Where(p => p.MemberType == PSMemberTypes.NoteProperty || p.MemberType == PSMemberTypes.Property)
            .Select(p => p.Name)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .ToList();

        if (properties.Count > 0)
        {
            var result = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in properties)
            {
                result[name] = ps.Properties[name]?.Value;
            }
            return result;
        }

        return ps.BaseObject;
    }
}
