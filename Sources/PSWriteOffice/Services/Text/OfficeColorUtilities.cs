using System;
using System.Management.Automation;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Text;

internal static class OfficeColorUtilities
{
    internal static string? ToRgbHex(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        try
        {
            return OfficeColor.Parse(color!).ToRgbHex();
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }
    }

    internal static PdfColor? ToPdfColor(string? color, string parameterName = "Color")
    {
        if (string.IsNullOrWhiteSpace(color))
        {
            return null;
        }

        OfficeColor parsed;
        try
        {
            parsed = OfficeColor.Parse(color!);
        }
        catch (FormatException exception)
        {
            throw new PSArgumentException(exception.Message, parameterName);
        }

        return PdfColor.FromRgb(parsed.R, parsed.G, parsed.B);
    }
}
