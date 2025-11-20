using ShapeCrawler;

namespace PSWriteOffice;

/// <summary>Extension helpers for ShapeCrawler slide collections.</summary>
public static class SlideCollectionExtensions
{
    /// <summary>Removes a slide at the specified index.</summary>
    public static void RemoveAt(this ISlideCollection slides, int index)
    {
        slides[index].Remove();
    }
}
