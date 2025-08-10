using System;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using HtmlToOpenXml;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static void AddHtml(WordDocument document, string html, HtmlImportMode mode = HtmlImportMode.Parse)
    {
        if (mode == HtmlImportMode.AsIs)
        {
            document.AddEmbeddedFragment(html, WordAlternativeFormatImportPartType.Html);
            return;
        }

        var field = typeof(WordDocument).GetField("_document", BindingFlags.NonPublic | BindingFlags.Instance)
                    ?? throw new InvalidOperationException("Unable to access underlying document.");
        var wordDocument = field.GetValue(document) as WordprocessingDocument
                           ?? throw new InvalidOperationException("Underlying document is null.");
        var mainPart = wordDocument.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");

        var converter = new HtmlConverter(mainPart);
        converter.ParseHtml(html);
    }
}
