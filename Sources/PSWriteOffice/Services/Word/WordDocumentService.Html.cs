using System;
using System.IO;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static void AddHtml(WordDocument document, string html, HtmlImportMode mode = HtmlImportMode.Parse)
    {
        var field = typeof(WordDocument).GetField("_document", BindingFlags.NonPublic | BindingFlags.Instance);
        if (field == null)
        {
            throw new InvalidOperationException("Unable to access underlying document.");
        }

        var wordDocument = field.GetValue(document) as WordprocessingDocument;
        if (wordDocument == null)
        {
            throw new InvalidOperationException("Underlying document is null.");
        }

        var mainPart = wordDocument.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing.");

        if (mode == HtmlImportMode.AsIs)
        {
            var altChunkId = "htmlChunk" + Guid.NewGuid().ToString("N");
            var part = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Html, altChunkId);
            using (var stream = part.GetStream())
            using (var writer = new StreamWriter(stream))
            {
                writer.Write(html);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            var body = mainPart.Document.Body ??= new Body();
            body.Append(altChunk);
            mainPart.Document.Save();
            return;
        }

        var converter = new HtmlConverter(mainPart);
        converter.ParseHtml(html);
    }
}
