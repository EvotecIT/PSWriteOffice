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
        ApplyAdditionalStyle(run, strike, color, null, fontSize, fontName);
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
        if (!string.IsNullOrWhiteSpace(spec.LinkUri) || !string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            AddHyperlink(paragraph, text, spec);
            return;
        }

        var underline = ResolveUnderline(spec.Underline, spec.UnderlineStyle);
        var run = paragraph.AddFormattedText(text, spec.Bold, spec.Italic, underline);
        ApplyAdditionalStyle(
            run,
            spec.Strike,
            spec.Color,
            spec.BackgroundColor,
            spec.FontSize.HasValue ? (int)Math.Round(spec.FontSize.Value) : null,
            spec.FontName);
    }

    private static void AddHyperlink(WordParagraph paragraph, string text, OfficeTextRunSpec spec)
    {
        if (!string.IsNullOrWhiteSpace(spec.LinkUri) && !string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            throw new PSArgumentException("A Word text run can target either LinkUri or LinkDestinationName, not both.", nameof(spec.LinkUri));
        }

        if (!string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            paragraph.AddHyperLink(text, spec.LinkDestinationName!, true, spec.LinkContents ?? string.Empty, true);
            return;
        }

        if (!Uri.TryCreate(spec.LinkUri, UriKind.Absolute, out var uri))
        {
            throw new PSArgumentException("Provide an absolute URL such as https://example.org or mailto:user@example.org.", nameof(spec.LinkUri));
        }

        paragraph.AddHyperLink(text, uri, true, spec.LinkContents ?? string.Empty, true);
    }

    private static void ApplyAdditionalStyle(WordParagraph run, bool strike, string? color, string? backgroundColor, int? fontSize, string? fontName)
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

        ApplyBackground(run, backgroundColor);

        if (fontSize.HasValue)
        {
            run.SetFontSize(fontSize.Value);
        }

        if (!string.IsNullOrWhiteSpace(fontName))
        {
            run.SetFontFamily(fontName!);
        }
    }

    private static void ApplyBackground(WordParagraph run, string? backgroundColor)
    {
        if (string.IsNullOrWhiteSpace(backgroundColor))
        {
            return;
        }

        if (OpenXmlValueParser.TryParse(backgroundColor, out HighlightColorValues highlight))
        {
            run.Highlight = highlight;
            return;
        }

        var rgb = OfficeColorUtilities.ToRgbHex(backgroundColor);
        if (!string.IsNullOrWhiteSpace(rgb))
        {
            run.ShadingFillColorHex = rgb!;
            run.ShadingPattern = ShadingPatternValues.Clear;
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
