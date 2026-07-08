using System;
using System.Management.Automation;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Services.Text;

namespace PSWriteOffice.Services.PowerPoint;

internal static class PowerPointTextRunService
{
    internal static void ApplyRuns(PowerPointTextBox textBox, object[] runs)
    {
        textBox.Text = string.Empty;
        var paragraph = textBox.Paragraphs.Count > 0 ? textBox.Paragraphs[0] : textBox.AddParagraph();
        ApplyRuns(paragraph, OfficeTextRunParser.ParseMany(runs), allowHyperlinks: true);
    }

    internal static void ApplyRuns(PowerPointTableCell cell, object[] runs)
    {
        cell.Text = string.Empty;
        var parsedRuns = OfficeTextRunParser.ParseMany(runs);
        var first = true;
        foreach (var spec in parsedRuns)
        {
            var run = first && cell.Runs.Count > 0 && string.IsNullOrEmpty(cell.Runs[0].Text)
                ? cell.Runs[0]
                : cell.AddRun(ToPowerPointText(spec));
            if (first && string.IsNullOrEmpty(run.Text))
            {
                run.Text = ToPowerPointText(spec);
            }

            ApplyStyle(run, spec, allowHyperlinks: false);
            first = false;
        }
    }

    private static void ApplyRuns(PowerPointParagraph paragraph, OfficeTextRunSpec[] runs, bool allowHyperlinks)
    {
        var first = true;
        foreach (var spec in runs)
        {
            var run = first && paragraph.Runs.Count > 0 && string.IsNullOrEmpty(paragraph.Runs[0].Text)
                ? paragraph.Runs[0]
                : paragraph.AddRun(ToPowerPointText(spec));
            if (first && string.IsNullOrEmpty(run.Text))
            {
                run.Text = ToPowerPointText(spec);
            }

            ApplyStyle(run, spec, allowHyperlinks);
            first = false;
        }
    }

    private static string ToPowerPointText(OfficeTextRunSpec run)
        => run.IsLineBreak ? Environment.NewLine : run.IsTab ? "\t" : run.Text;

    private static void ApplyStyle(PowerPointTextRun run, OfficeTextRunSpec spec, bool allowHyperlinks)
    {
        run.Bold = spec.Bold;
        run.Italic = spec.Italic;
        run.Underline = spec.Underline;
        run.Strikethrough = spec.Strike;
        run.FontSize = spec.FontSize.HasValue
            ? (int)Math.Round(spec.FontSize.Value)
            : null;
        run.FontName = !string.IsNullOrWhiteSpace(spec.FontName)
            ? spec.FontName
            : null;
        run.Color = OfficeColorUtilities.ToRgbHex(spec.Color);
        run.HighlightColor = OfficeColorUtilities.ToRgbHex(spec.BackgroundColor);
        run.ClearHyperlink();

        if (!string.IsNullOrWhiteSpace(spec.LinkDestinationName))
        {
            throw new PSArgumentException("PowerPoint text runs support URI links, not named PDF/Word destinations.", nameof(spec.LinkDestinationName));
        }

        if (!string.IsNullOrWhiteSpace(spec.LinkUri))
        {
            if (!allowHyperlinks)
            {
                throw new PSArgumentException("PowerPoint table cell runs do not support hyperlinks yet.", nameof(spec.LinkUri));
            }

            run.SetHyperlink(spec.LinkUri!, spec.LinkContents);
        }
    }
}
