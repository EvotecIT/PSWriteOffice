using System;
using System.Drawing;
using ClosedXML.Excel;

namespace PSWriteOffice.Services;

public static class ColorService
{
    public static XLColor? GetColor(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
            return null;

        input = input!.Trim();
        try
        {
            // Try HTML notation
            return XLColor.FromHtml(input);
        }
        catch
        {
            // ignore
        }

        try
        {
            // Try known color names
            return XLColor.FromName(input);
        }
        catch
        {
            // ignore
        }

        try
        {
            // Try RGB "r,g,b" format
            var parts = input.Split(',');
            if (parts.Length == 3)
            {
                var r = byte.Parse(parts[0]);
                var g = byte.Parse(parts[1]);
                var b = byte.Parse(parts[2]);
                return XLColor.FromColor(Color.FromArgb(r, g, b));
            }
        }
        catch
        {
            // ignore
        }

        return null;
    }
}
