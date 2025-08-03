using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static void RemoveFooters(WordDocument document)
    {
        WordFooter.RemoveFooters(document);
    }

    public static void RemoveHeaders(WordDocument document)
    {
        WordHeader.RemoveHeaders(document);
    }
}
