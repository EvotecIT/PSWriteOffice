using System;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using HtmlToOpenXml;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    /// <summary>Adds HTML content to the document.</summary>
    public static void AddHtml(WordDocument document, string html, HtmlImportMode mode = HtmlImportMode.Parse)
    {
        if (mode == HtmlImportMode.AsIs)
        {
            // For AsIs mode, embed the HTML directly without parsing
            document.AddEmbeddedFragment(html, WordAlternativeFormatImportPartType.Html);
            return;
        }

        // For Parse mode, use HtmlToOpenXml to convert HTML to Word elements
        // The _wordprocessingDocument field is public in OfficeIMO
        var field = typeof(WordDocument).GetField("_wordprocessingDocument");
        if (field == null)
        {
            throw new InvalidOperationException("Unable to access underlying document field.");
        }

        var wordDocument = field.GetValue(document) as WordprocessingDocument;
        if (wordDocument == null)
        {
            throw new InvalidOperationException("Underlying document is null.");
        }

        var mainPart = wordDocument.MainDocumentPart;
        if (mainPart == null)
        {
            throw new InvalidOperationException("Main document part is missing.");
        }

        var converter = new HtmlConverter(mainPart);
        var body = mainPart.Document?.Body ?? throw new InvalidOperationException("Document body is missing.");
#pragma warning disable CS0618
        converter.ParseHtml(html);
#pragma warning restore CS0618
    }
}
