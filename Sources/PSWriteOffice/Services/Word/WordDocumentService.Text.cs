using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static WordParagraph AddText(WordDocument? document, WordParagraph? paragraph, string[] text, bool?[]? bold,
        bool?[]? italic, UnderlineValues?[]? underline, string[]? color, JustificationValues? alignment,
        WordParagraphStyles? style)
    {
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
