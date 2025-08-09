using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static WordParagraph AddText(WordDocument? document, WordParagraph? paragraph, string[] text, bool?[]? bold,
        bool?[]? italic, UnderlineValues?[]? underline, string[]? color, JustificationValues? alignment,
        WordParagraphStyles? style)
    {
        if (bold != null && bold.Length != text.Length)
        {
            throw new ArgumentException("bold length must match text length", nameof(bold));
        }
        if (italic != null && italic.Length != text.Length)
        {
            throw new ArgumentException("italic length must match text length", nameof(italic));
        }
        if (underline != null && underline.Length != text.Length)
        {
            throw new ArgumentException("underline length must match text length", nameof(underline));
        }
        if (color != null && color.Length != text.Length)
        {
            throw new ArgumentException("color length must match text length", nameof(color));
        }

        var para = paragraph ?? document!.AddParagraph();

        for (var t = 0; t < text.Length; t++)
        {
            para = para.AddText(text[t]);

            if (bold != null && t < bold.Length && bold[t].HasValue)
            {
                para.Bold = bold[t].Value;
            }
            if (italic != null && t < italic.Length && italic[t].HasValue)
            {
                para.Italic = italic[t].Value;
            }
        }

        return para;
    }
}
