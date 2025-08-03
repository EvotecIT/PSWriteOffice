using System;
using System.Collections;
using OfficeIMO.Word;

namespace PSWriteOffice.Services.Word;

public static partial class WordDocumentService
{
    public static WordList AddList(WordDocument document, WordListStyle style)
    {
        return document.AddList(style);
    }

    public static object AddListItem(WordList list, int level, string[] text)
    {
        var method = typeof(WordList).GetMethod("AddItem", new[] { typeof(string[]), typeof(int) });
        if (method != null)
        {
            return method.Invoke(list, new object[] { text, level })!;
        }

        method = typeof(WordList).GetMethod("AddItem", new[] { typeof(string), typeof(int) });
        if (method != null)
        {
            return method.Invoke(list, new object[] { string.Join(" ", text), level })!;
        }

        throw new InvalidOperationException("AddItem method not found.");
    }
}
