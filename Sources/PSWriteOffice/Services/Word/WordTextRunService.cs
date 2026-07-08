using System;
using System.Management.Automation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.Word;

internal static class WordTextRunService
{
    internal static void ApplyRuns(WordParagraph paragraph, object[] runs)
        => ApplyRuns(paragraph, OfficeTextRunParser.ParseMany(runs));

    internal static void ApplyRuns(WordParagraph paragraph, OfficeTextRunSpec[] runs)
    {
        foreach (var run in runs)
        {
            ApplyRun(paragraph, run);
        }
    }

    internal static WordParagraph AddText(
        WordParagraph paragraph,
        string text,
        bool bold,
        bool italic,
        UnderlineValues? underline,
        bool strike,
        string? color,
        int? fontSize,
        string? fontName)
    {
        var run = paragraph.AddFormattedText(text, bold, italic, underline);
        ApplyAdditionalStyle(run, strike, color, fontSize, fontName);
        return run;
    }

    private static void ApplyRun(WordParagraph paragraph, OfficeTextRunSpec spec)
    {
        if (spec.IsLineBreak)
        {
            paragraph.AddBreak();
            return;
        }

        var text = spec.IsTab ? "\t" : spec.Text;
        var underline = ResolveUnderline(spec.Underline, spec.UnderlineStyle);
        var run = paragraph.AddFormattedText(text, spec.Bold, spec.Italic, underline);
        ApplyAdditionalStyle(
            run,
            spec.Strike,
            spec.Color,
            spec.FontSize.HasValue ? (int)Math.Round(spec.FontSize.Value) : null,
            spec.FontName);
    }

    private static void ApplyAdditionalStyle(WordParagraph run, bool strike, string? color, int? fontSize, string? fontName)
    {
        if (strike)
        {
            run.SetStrike();
        }

        var rgb = OfficeColorUtilities.ToRgbHex(color);
        if (!string.IsNullOrWhiteSpace(rgb))
        {
            run.SetColorHex(rgb!);
        }

        if (fontSize.HasValue)
        {
            run.SetFontSize(fontSize.Value);
        }

        if (!string.IsNullOrWhiteSpace(fontName))
        {
            run.SetFontFamily(fontName!);
        }
    }

    private static UnderlineValues? ResolveUnderline(bool underline, string? underlineStyle)
    {
        if (!underline)
        {
            return null;
        }

        if (string.IsNullOrWhiteSpace(underlineStyle))
        {
            return UnderlineValues.Single;
        }

        return OpenXmlValueParser.TryParse<UnderlineValues>(underlineStyle, out var parsed)
            ? parsed
            : throw new PSArgumentException($"Unsupported Word underline style '{underlineStyle}'.", nameof(underlineStyle));
    }
}
