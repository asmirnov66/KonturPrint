using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KonturPrint.Extensions
{
    public static class OpenXmlElementExtensions
    {
        public static T GetFirstDescendant<T>(this OpenXmlElement parent) where T : OpenXmlElement
        {
            var descendants = parent.Descendants<T>();
            return descendants?.FirstOrDefault();
        }

        public static T GetParent<T>(this OpenXmlElement child) where T : OpenXmlElement
        {
            while (child != null)
            {
                child = child.Parent;
                if (child is T)
                    return (T)child;
            }
            return null;
        }

        public static bool IsEndBookmark(this OpenXmlElement element, BookmarkStart startBookmark)
        {
            return IsEndBookmark(element as BookmarkEnd, startBookmark);
        }

        public static bool IsEndBookmark(this BookmarkEnd endBookmark, BookmarkStart startBookmark)
        {
            if (endBookmark == null)
                return false;
            return endBookmark.Id == startBookmark.Id;
        }
    }
}
