using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Management.Automation;

namespace PSWriteOffice.Services.Text;

internal static class OfficeTextRunParser
{
    internal static OfficeTextRunSpec[] ParseMany(object[]? runs)
    {
        var normalized = ToRunArray(runs);
        if (normalized.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text run.");
        }

        return normalized.Select(Parse).ToArray();
    }

    internal static OfficeTextRunSpec[] ParseMany(object? runs)
    {
        var normalized = ToRunArray(runs);
        if (normalized.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text run.");
        }

        return normalized.Select(Parse).ToArray();
    }

    internal static object[] ToRunArray(object? runs)
    {
        if (runs == null)
        {
            return Array.Empty<object>();
        }

        if (runs is string)
        {
            return new[] { runs };
        }

        if (runs is IDictionary || runs is OfficeTextRunSpec)
        {
            return new[] { runs };
        }

        return runs is IEnumerable enumerable
            ? enumerable.Cast<object>().ToArray()
            : new[] { runs };
    }

    internal static OfficeTextRunSpec Parse(object value)
    {
        if (value is OfficeTextRunSpec spec)
        {
            return spec;
        }

        if (value is string text)
        {
            return new OfficeTextRunSpec { Text = text };
        }

        var kind = GetString(value, "Type", "Kind", "Run");
        var normalizedKind = NormalizeKind(kind);
        var underline = GetUnderline(value, out var underlineStyle) ||
                        normalizedKind is "underline" or "underlined" or "link" or "bookmarklink";

        var baseline = GetString(value, "Baseline");
        if (normalizedKind == "superscript")
        {
            baseline = "Superscript";
        }
        else if (normalizedKind == "subscript")
        {
            baseline = "Subscript";
        }

        return new OfficeTextRunSpec
        {
            Text = GetString(value, "Text", "Value", "Content") ?? string.Empty,
            Kind = kind,
            Bold = GetBool(value, "Bold") || normalizedKind == "bold",
            Italic = GetBool(value, "Italic") || normalizedKind == "italic",
            Underline = underline,
            UnderlineStyle = underlineStyle,
            Strike = GetBool(value, "Strike", "Strikethrough") || normalizedKind is "strike" or "strikethrough",
            Color = GetString(value, "Color", "TextColor", "FontColor"),
            BackgroundColor = GetString(value, "BackgroundColor", "HighlightColor", "FillColor"),
            FontSize = GetDouble(value, "FontSize", "Size"),
            FontName = GetString(value, "FontName", "Font", "Typeface", "FontFamily"),
            Baseline = baseline,
            LinkUri = GetString(value, "LinkUri", "Uri", "Url", "Href"),
            LinkDestinationName = GetString(value, "LinkDestinationName", "DestinationName", "Bookmark", "BookmarkName"),
            LinkContents = GetString(value, "LinkContents", "Contents", "Tooltip"),
            TabLeader = GetString(value, "Leader", "TabLeader"),
            TabAlignment = GetString(value, "Alignment", "TabAlignment")
        };
    }

    internal static string GetPlainText(OfficeTextRunSpec[] runs)
        => string.Concat(runs.Select(run => run.IsLineBreak ? Environment.NewLine : run.IsTab ? "\t" : run.Text));

    internal static string NormalizeKind(string? value)
        => (value ?? string.Empty).Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();

    internal static string? GetString(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToString(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static double? GetDouble(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToDouble(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static bool GetBool(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value != null && Convert.ToBoolean(UnwrapPSObject(value), CultureInfo.InvariantCulture);
    }

    internal static object? GetValue(object? source, params string[] names)
    {
        source = UnwrapPSObject(source);
        if (source is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (names.Any(name => string.Equals(name, key, StringComparison.OrdinalIgnoreCase)))
                {
                    return entry.Value;
                }
            }

            return null;
        }

        if (source is null)
        {
            return null;
        }

        var psObject = PSObject.AsPSObject(source);
        foreach (var name in names)
        {
            var property = psObject.Properties
                .Cast<PSPropertyInfo>()
                .FirstOrDefault(candidate => candidate.IsGettable && string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
            {
                return property.Value;
            }
        }

        return null;
    }

    private static bool GetUnderline(object source, out string? underlineStyle)
    {
        underlineStyle = null;
        var underlineValue = GetValue(source, "Underline", "Underlined");
        var styleValue = GetValue(source, "UnderlineStyle", "UnderlineKind");
        if (styleValue != null)
        {
            underlineStyle = Convert.ToString(UnwrapPSObject(styleValue), CultureInfo.InvariantCulture);
            return !IsFalseUnderline(underlineStyle);
        }

        if (underlineValue == null)
        {
            return false;
        }

        underlineValue = UnwrapPSObject(underlineValue);
        if (underlineValue is bool boolValue)
        {
            return boolValue;
        }

        underlineStyle = Convert.ToString(underlineValue, CultureInfo.InvariantCulture);
        return !IsFalseUnderline(underlineStyle);
    }

    private static bool IsFalseUnderline(string? value)
    {
        var normalized = NormalizeKind(value);
        return normalized is "" or "false" or "none" or "no" or "off";
    }

    private static object? UnwrapPSObject(object? value)
        => value is PSObject psObject ? psObject.BaseObject : value;
}
