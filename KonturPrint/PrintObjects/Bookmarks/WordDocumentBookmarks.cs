using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.Bookmarks
{
    public class WordDocumentBookmarks : IWordDocumentBookmarks
    {
        private WordprocessingDocument Doc { get; }
        private IDictionary<string, IWordDocumentBookmark> Items { get; set; }
        private IWordDocumentStructure WordDocument { get; set; }

        public int Count
        {
            get
            {
                FillItems();
                return Items.Count;
            }
        }

        public WordDocumentBookmarks(WordprocessingDocument doc)
        {
            Doc = doc;
            Items = new Dictionary<string, IWordDocumentBookmark>(StringComparer.OrdinalIgnoreCase);
        }

        public WordDocumentBookmarks(IWordDocumentStructure wordDoc)
        {
            Doc = wordDoc.InnerDoc;
            WordDocument = wordDoc;
            Items = new Dictionary<string, IWordDocumentBookmark>(StringComparer.OrdinalIgnoreCase);
        }

        public IWordDocumentBookmark Item(string index)
        {
            IWordDocumentBookmark bkm;
            if (TryGetBookmark(index, out bkm))
            {
                return bkm;
            }
            return null;
        }

        public IWordDocumentBookmark Item(int index)
        {
            if (index <= 0)
            {
                return null;
            }
            var bn = index - 1;
            if (bn < Items.Count)
            {
                return Items.ElementAt(bn).Value;
            }
            return null;
        }

        public bool Exists(string name)
        {
            IWordDocumentBookmark bkm;
            if (TryGetBookmark(name, out bkm))
            {
                return true;
            }
            return false;
        }

        private void FillItems()
        {
            Items = new Dictionary<string, IWordDocumentBookmark>(StringComparer.OrdinalIgnoreCase);
            FillItemsFromMain();
            FillItemsFromHeaders();
            FillItemsFromFooters();
        }

        private void FillItemsFromMain()
        {
            var main = Doc.MainDocumentPart;
            foreach (var bkm in main.Document.Body.Descendants<BookmarkStart>())
            {
                var newBkm = new WordDocumentBookmark(Doc);
                if (newBkm.Select(bkm.Name))
                {
                    Items.Add(newBkm.Name, newBkm);
                }
            }
        }

        private void FillItemsFromHeaders()
        {
            if (Doc.MainDocumentPart.HeaderParts != null)
            {
                foreach (var header in Doc.MainDocumentPart.HeaderParts)
                {
                    foreach (var bkm in header.RootElement.Descendants<BookmarkStart>())
                    {
                        var newBkm = new WordDocumentBookmark(Doc);
                        if (newBkm.Select(bkm.Name))
                        {
                            Items.Add(newBkm.Name, newBkm);
                        }
                    }
                }
            }
        }

        private void FillItemsFromFooters()
        {
            if (Doc.MainDocumentPart.FooterParts != null)
            {
                foreach (var footer in Doc.MainDocumentPart.FooterParts)
                {
                    foreach (var bkm in footer.RootElement.Descendants<BookmarkStart>())
                    {
                        var newBkm = new WordDocumentBookmark(Doc);
                        if (newBkm.Select(bkm.Name))
                        {
                            Items.Add(newBkm.Name, newBkm);
                        }
                    }
                }
            }
        }

        private bool TryGetBookmark(string name, out IWordDocumentBookmark bookmark)
        {
            IWordDocumentBookmark bkm;
            if (Items.TryGetValue(name, out bkm))
            {
                bookmark = bkm;
                return true;
            }
            bkm = GetBookmark(name);
            if (bkm != null)
            {
                Items.Add(bkm.Name, bkm);
                bookmark = bkm;
                return true;
            }
            bookmark = null;
            return false;
        }

        private IWordDocumentBookmark GetBookmark(string name)
        {
            var bkm = new WordDocumentBookmark(Doc);
            if (bkm.Select(name))
            {
                return bkm;
            }
            return null;
        }
    }
}
