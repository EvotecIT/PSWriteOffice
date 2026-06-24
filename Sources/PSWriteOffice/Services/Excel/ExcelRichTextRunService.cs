using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelRichTextRunService
{
    public static ExcelRichTextRun[] ToRuns(object[] runs)
    {
        if (runs.Length == 0)
        {
            throw new PSArgumentException("Provide at least one Excel rich text run.");
        }

        return runs.Select(ToRun).ToArray();
    }

    public static PSObject CreateRecord(
        ExcelRichTextRun run,
        int index,
        string address,
        int row,
        int column,
        string sheetName,
        string? path)
    {
        var record = new PSObject();
        record.Properties.Add(new PSNoteProperty("Index", index));
        record.Properties.Add(new PSNoteProperty("Text", run.Text));
        record.Properties.Add(new PSNoteProperty("Bold", run.Bold));
        record.Properties.Add(new PSNoteProperty("Italic", run.Italic));
        record.Properties.Add(new PSNoteProperty("Underline", run.Underline));
        record.Properties.Add(new PSNoteProperty("FontColor", run.FontColor));
        record.Properties.Add(new PSNoteProperty("Color", run.FontColor));
        record.Properties.Add(new PSNoteProperty("FontName", run.FontName));
        record.Properties.Add(new PSNoteProperty("FontSize", run.FontSize));
        record.Properties.Add(new PSNoteProperty("Address", address));
        record.Properties.Add(new PSNoteProperty("Row", row));
        record.Properties.Add(new PSNoteProperty("Column", column));
        record.Properties.Add(new PSNoteProperty("SheetName", sheetName));
        record.Properties.Add(new PSNoteProperty("Sheet", sheetName));
        if (!string.IsNullOrWhiteSpace(path))
        {
            record.Properties.Add(new PSNoteProperty("Path", path));
            record.Properties.Add(new PSNoteProperty("InputPath", path));
        }

        return record;
    }

    private static ExcelRichTextRun ToRun(object value)
    {
        if (value is ExcelRichTextRun richTextRun)
        {
            return richTextRun;
        }

        if (value is string text)
        {
            return new ExcelRichTextRun(text);
        }

        return new ExcelRichTextRun(GetString(value, "Text", "Value") ?? string.Empty)
        {
            Bold = GetBool(value, "Bold"),
            Italic = GetBool(value, "Italic"),
            Underline = GetBool(value, "Underline", "Underlined"),
            FontColor = GetString(value, "FontColor", "Color", "TextColor"),
            FontName = GetString(value, "FontName", "Font", "Typeface"),
            FontSize = GetDouble(value, "FontSize", "Size")
        };
    }

    private static string? GetString(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static bool GetBool(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value != null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    private static double? GetDouble(object source, params string[] names)
    {
        var value = GetValue(source, names);
        return value == null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
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
}
