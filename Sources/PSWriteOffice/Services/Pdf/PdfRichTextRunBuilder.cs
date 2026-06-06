using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfRichTextRunBuilder
{
    internal static void ApplyText(
        PdfParagraphBuilder builder,
        string[] text,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        PdfColor? color,
        PdfColor? backgroundColor,
        double? fontSize,
        PdfStandardFont? font,
        string? linkUri,
        string? linkDestinationName,
        string? linkContents)
    {
        if (text.Length == 0)
        {
            throw new PSArgumentException("Provide at least one text value.");
        }

        if (linkUri != null && linkDestinationName != null)
        {
            throw new PSArgumentException("A PDF text link can target either -LinkUri or -LinkDestinationName, not both.");
        }

        ApplyStyle(builder, bold, italic, underline, strike, baseline, color, backgroundColor, fontSize, font);
        foreach (var value in text)
        {
            AddText(builder, value, linkUri, linkDestinationName, linkContents, color, underline || linkUri != null || linkDestinationName != null);
        }
    }

    internal static void ApplyRuns(PdfParagraphBuilder builder, object[] runs)
    {
        if (runs.Length == 0)
        {
            throw new PSArgumentException("Provide at least one PDF text run.");
        }

        foreach (var run in runs)
        {
            ApplyRun(builder, run);
        }
    }

    internal static void ApplyRuns(PdfParagraphBuilder builder, object? runs)
    {
        ApplyRuns(builder, ToRunArray(runs));
    }

    internal static object[] ToRunArray(object? runs)
    {
        if (runs == null)
        {
            return Array.Empty<object>();
        }

        if (runs is string text)
        {
            return new object[] { text };
        }

        if (runs is IDictionary)
        {
            return new[] { runs };
        }

        return runs is IEnumerable enumerable
            ? enumerable.Cast<object>().ToArray()
            : new[] { runs };
    }

    private static void ApplyRun(PdfParagraphBuilder builder, object run)
    {
        if (run is string literalText)
        {
            ResetStyle(builder);
            builder.Text(literalText);
            return;
        }

        var type = Normalize(GetString(run, "Type", "Kind", "Run") ?? string.Empty);
        if (type == "linebreak" || type == "break" || type == "br")
        {
            builder.LineBreak();
            return;
        }

        if (type == "tab")
        {
            builder.Tab(GetEnum(run, PdfTabLeaderStyle.None, "Leader", "TabLeader"), GetEnum(run, PdfTabAlignment.Left, "Alignment", "TabAlignment"));
            return;
        }

        var bold = GetBool(run, "Bold") || type == "bold";
        var italic = GetBool(run, "Italic") || type == "italic";
        var underline = GetBool(run, "Underline", "Underlined") || type == "underline" || type == "underlined" || type == "link" || type == "bookmarklink";
        var strike = GetBool(run, "Strike", "Strikethrough") || type == "strike" || type == "strikethrough";
        var baseline = GetEnum(run, PdfTextBaseline.Normal, "Baseline");
        if (type == "superscript")
        {
            baseline = PdfTextBaseline.Superscript;
        }
        else if (type == "subscript")
        {
            baseline = PdfTextBaseline.Subscript;
        }

        var color = PdfCommandUtilities.ParseColor(GetString(run, "Color", "TextColor"));
        var backgroundColor = PdfCommandUtilities.ParseColor(GetString(run, "BackgroundColor", "HighlightColor"));
        var fontSize = GetDouble(run, "FontSize");
        var font = GetNullableEnum<PdfStandardFont>(run, "Font");
        var linkUri = GetString(run, "LinkUri", "Uri", "Url", "Href");
        var linkDestinationName = GetString(run, "LinkDestinationName", "DestinationName", "Bookmark", "BookmarkName");
        var linkContents = GetString(run, "LinkContents", "Contents", "Tooltip");
        var text = GetString(run, "Text", "Value") ?? string.Empty;

        if (linkUri != null && linkDestinationName != null)
        {
            throw new PSArgumentException("A PDF text run can target either LinkUri or LinkDestinationName, not both.");
        }

        ApplyStyle(builder, bold, italic, underline, strike, baseline, color, backgroundColor, fontSize, font);
        AddText(builder, text, linkUri, linkDestinationName, linkContents, color, underline || linkUri != null || linkDestinationName != null);
    }

    private static void AddText(PdfParagraphBuilder builder, string text, string? linkUri, string? linkDestinationName, string? linkContents, PdfColor? color, bool underline)
    {
        if (!string.IsNullOrWhiteSpace(linkUri))
        {
            builder.Link(text, linkUri!, color, underline, linkContents);
            return;
        }

        if (!string.IsNullOrWhiteSpace(linkDestinationName))
        {
            builder.LinkToBookmark(text, linkDestinationName!, color, underline, linkContents);
            return;
        }

        builder.Text(text);
    }

    private static void ApplyStyle(
        PdfParagraphBuilder builder,
        bool bold,
        bool italic,
        bool underline,
        bool strike,
        PdfTextBaseline baseline,
        PdfColor? color,
        PdfColor? backgroundColor,
        double? fontSize,
        PdfStandardFont? font)
    {
        builder.Bold(bold)
            .Italic(italic)
            .Underline(underline)
            .Strike(strike)
            .Baseline(baseline);

        if (color.HasValue)
        {
            builder.Color(color.Value);
        }
        else
        {
            builder.ResetColor();
        }

        if (backgroundColor.HasValue)
        {
            builder.BackgroundColor(backgroundColor.Value);
        }
        else
        {
            builder.ResetBackgroundColor();
        }

        if (fontSize.HasValue)
        {
            builder.FontSize(fontSize.Value);
        }
        else
        {
            builder.ResetFontSize();
        }

        if (font.HasValue)
        {
            builder.Font(font.Value);
        }
        else
        {
            builder.ResetFont();
        }
    }

    private static void ResetStyle(PdfParagraphBuilder builder)
    {
        ApplyStyle(builder, bold: false, italic: false, underline: false, strike: false, PdfTextBaseline.Normal, null, null, null, null);
    }

    private static string? GetString(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static double? GetDouble(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    private static bool GetBool(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value != null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    private static TEnum GetEnum<TEnum>(object source, TEnum defaultValue, params string[] names)
        where TEnum : struct
    {
        return GetNullableEnum<TEnum>(source, names) ?? defaultValue;
    }

    private static TEnum? GetNullableEnum<TEnum>(object source, params string[] names)
        where TEnum : struct
    {
        var value = GetValue(source, names);
        if (value == null)
        {
            return null;
        }

        return value is TEnum enumValue
            ? enumValue
            : (TEnum)Enum.Parse(typeof(TEnum), Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty, ignoreCase: true);
    }

    private static object? GetValue(object source, params string[] names)
    {
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
        }

        var psObject = PSObject.AsPSObject(source);
        foreach (var name in names)
        {
            var property = psObject.Properties
                .Cast<PSPropertyInfo>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
            {
                return property.Value;
            }
        }

        return null;
    }

    private static string Normalize(string value)
    {
        return value.Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
    }
}
