using System;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

internal static class WordHostExtensions
{
    public static WordParagraph AddParagraph(this object host, string? text = null)
    {
        switch (host)
        {
            case WordSection section:
                return string.IsNullOrEmpty(text) ? section.AddParagraph() : section.AddParagraph(text);
            case WordHeader header:
                return string.IsNullOrEmpty(text) ? header.AddParagraph() : header.AddParagraph(text!);
            case WordFooter footer:
                return string.IsNullOrEmpty(text) ? footer.AddParagraph() : footer.AddParagraph(text!);
            default:
                throw new InvalidOperationException("Paragraphs can only be added inside sections, headers, or footers.");
        }
    }

    public static WordList AddList(this WordDocument document, WordListStyle style)
    {
        return document.AddList(style);
    }
}
