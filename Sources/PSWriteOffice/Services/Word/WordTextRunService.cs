using System;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using System.Reflection;
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
        var underline = ResolveUnderline(spec.Underline, spec.UnderlineStyle);
        if (!string.IsNullOrWhiteSpace(spec.LinkUri) || !string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            AddHyperlink(paragraph, text, spec);
            ApplyHyperlinkStyle(paragraph.Hyperlink, spec, underline);
            return;
        }

        var run = paragraph.AddFormattedText(text, spec.Bold, spec.Italic, underline);
        ApplyAdditionalStyle(
            run,
            spec.Strike,
            spec.Color,
            spec.BackgroundColor,
            spec.FontSize.HasValue ? (int)Math.Round(spec.FontSize.Value) : null,
            spec.FontName);
        ApplyBaseline(run, spec.Baseline);
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

    private static void ApplyHyperlinkStyle(WordHyperLink? hyperlink, OfficeTextRunSpec spec, UnderlineValues? underline)
    {
        var run = GetHyperlinkRun(hyperlink);
        if (run == null)
        {
            return;
        }

        var properties = EnsureRunProperties(run);
        if (spec.Bold)
        {
            properties.Bold = new Bold();
        }

        if (spec.Italic)
        {
            properties.Italic = new Italic();
        }

        if (underline.HasValue)
        {
            properties.Underline = new Underline { Val = underline.Value };
        }

        if (spec.Strike)
        {
            properties.Strike = new Strike();
        }

        var rgb = OfficeColorUtilities.ToRgbHex(spec.Color);
        if (!string.IsNullOrWhiteSpace(rgb))
        {
            properties.Color = new Color { Val = rgb! };
        }

        ApplyBackground(properties, spec.BackgroundColor);

        if (spec.FontSize.HasValue)
        {
            var halfPoints = ((int)Math.Round(spec.FontSize.Value) * 2).ToString(CultureInfo.InvariantCulture);
            properties.FontSize = new FontSize { Val = halfPoints };
        }

        if (!string.IsNullOrWhiteSpace(spec.FontName))
        {
            properties.RunFonts = new RunFonts
            {
                Ascii = spec.FontName,
                HighAnsi = spec.FontName,
                EastAsia = spec.FontName,
                ComplexScript = spec.FontName
            };
        }

        ApplyBaseline(properties, spec.Baseline);
    }

    private static Run? GetHyperlinkRun(WordHyperLink? hyperlink)
    {
        if (hyperlink == null)
        {
            return null;
        }

        // OfficeIMO exposes hyperlink metadata publicly, but not the formatted inner run.
        var field = typeof(WordHyperLink).GetField("_hyperlink", BindingFlags.Instance | BindingFlags.NonPublic);
        return field?.GetValue(hyperlink) is Hyperlink element
            ? element.Elements<Run>().FirstOrDefault()
            : null;
    }

    private static RunProperties EnsureRunProperties(Run run)
    {
        if (run.RunProperties != null)
        {
            return run.RunProperties;
        }

        var properties = new RunProperties();
        run.PrependChild(properties);
        return properties;
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

    private static void ApplyBaseline(WordParagraph run, string? baseline)
    {
        var value = ResolveBaseline(baseline);
        if (value == null)
        {
            return;
        }

        if (value == VerticalPositionValues.Superscript)
        {
            run.SetSuperScript();
            return;
        }

        if (value == VerticalPositionValues.Subscript)
        {
            run.SetSubScript();
        }
    }

    private static void ApplyBaseline(RunProperties properties, string? baseline)
    {
        var value = ResolveBaseline(baseline);
        if (value.HasValue)
        {
            properties.VerticalTextAlignment = new VerticalTextAlignment { Val = value.Value };
        }
    }

    private static VerticalPositionValues? ResolveBaseline(string? baseline)
    {
        if (string.IsNullOrWhiteSpace(baseline))
        {
            return null;
        }

        switch (OfficeTextRunParser.NormalizeKind(baseline))
        {
            case "normal":
            case "baseline":
                return null;
            case "superscript":
            case "super":
            case "sup":
                return VerticalPositionValues.Superscript;
            case "subscript":
            case "sub":
                return VerticalPositionValues.Subscript;
            default:
                throw new PSArgumentException($"Unsupported Word baseline '{baseline}'.", nameof(baseline));
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

    private static void ApplyBackground(RunProperties properties, string? backgroundColor)
    {
        if (string.IsNullOrWhiteSpace(backgroundColor))
        {
            return;
        }

        if (OpenXmlValueParser.TryParse(backgroundColor, out HighlightColorValues highlight))
        {
            properties.Highlight = new Highlight { Val = highlight };
            return;
        }

        var rgb = OfficeColorUtilities.ToRgbHex(backgroundColor);
        if (!string.IsNullOrWhiteSpace(rgb))
        {
            properties.Shading = new Shading { Fill = rgb!, Val = ShadingPatternValues.Clear };
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
