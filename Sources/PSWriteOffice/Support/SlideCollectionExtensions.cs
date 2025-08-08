using ShapeCrawler;

namespace PSWriteOffice;

public static class SlideCollectionExtensions
{
    public static void RemoveAt(this ISlideCollection slides, int index)
    {
        slides[index].Remove();
    }
}
