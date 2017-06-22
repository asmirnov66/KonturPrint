using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Extensions;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.Tables;

namespace KonturPrint.PrintObjects.Bookmarks
{
    public class WordDocumentBookmark : IWordDocumentBookmark, IPrintObject
    {
        private WordprocessingDocument Doc { get; }
        private BookmarkStart BkmStart { get; set; }
        private BookmarkEnd BkmEnd { get; set; }
        private OpenXmlElement SiblingElement { get; set; }

        public string Name { get; private set; }

        public string Text
        {
            get { return GetText(); }
            set { SetText(value); }
        }

        public IWordDocumentTable Table { get; private set; }
        public OpenXmlElement XmlElement { get; private set; }

        public WordDocumentBookmark(WordprocessingDocument doc)
        {
            Doc = doc;
            BkmStart = null;
            BkmEnd = null;
        }

        public bool Select(string name)
        {
            var bkmStart = FindBookmarkStartInMain(name);
            if (bkmStart != null)
            {
                BkmStart = bkmStart;
                Name = BkmStart.Name;
                BkmEnd = FindBookmarkEndInMain(bkmStart.Id);
                TryFindTable();
                XmlElement = BkmStart;
                return true;
            }

            var bkmHeader = FindBookmarkStartInHeaders(name);
            if (bkmHeader != null)
            {
                BkmStart = bkmHeader.Item1;
                var header = bkmHeader.Item2;
                Name = BkmStart.Name;
                BkmEnd = FindBookmarkEndInHeader(header, bkmHeader.Item1.Id);
                TryFindTable();
                XmlElement = BkmStart;
                return true;
            }

            var bkmFooter = FindBookmarkStartInFooters(name);
            if (bkmFooter != null)
            {
                BkmStart = bkmFooter.Item1;
                var footer = bkmFooter.Item2;
                Name = BkmStart.Name;
                BkmEnd = FindBookmarkEndInFooter(footer, bkmFooter.Item1.Id);
                TryFindTable();
                XmlElement = BkmStart;
                return true;
            }
            return false;
        }

        public IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            var copy = GetBookmarkCopy();
            var destObject = destPrintObject.XmlElement.AppendChild(copy.Item1);
            destPrintObject.XmlElement.AppendChild(copy.Item2);
            return new PrintObject(destObject);
        }

        public IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            if (TryCopyTable(sourcePrintObject))
            {
                return new PrintObject((Table)Table.Table);
            }
            var sourceObject = (OpenXmlElement)sourcePrintObject.XmlElement.Clone();
            var destObject = BkmStart.Parent.InsertAfter(sourceObject, BkmStart);
            return new PrintObject(destObject);
        }

        private BookmarkStart FindBookmarkStartInMain(string name)
        {
            var main = Doc.MainDocumentPart;
            var bkmStart = main.Document.Body.Descendants<BookmarkStart>()
                .FirstOrDefault(b => string.Equals(b.Name.ToString(), name, StringComparison.OrdinalIgnoreCase));
            return bkmStart;
        }

        private BookmarkEnd FindBookmarkEndInMain(string idBkm)
        {
            var main = Doc.MainDocumentPart;
            var bkmEnd = main.Document.Body.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => string.Equals(b.Id, idBkm, StringComparison.OrdinalIgnoreCase));
            return bkmEnd;
        }

        private Tuple<BookmarkStart, HeaderPart> FindBookmarkStartInHeaders(string name)
        {
            if (Doc.MainDocumentPart.HeaderParts != null)
            {
                foreach (var header in Doc.MainDocumentPart.HeaderParts)
                {
                    var bkmStart = FindBookmarkStartInHeader(header, name);
                    if (bkmStart != null)
                    {
                        return new Tuple<BookmarkStart, HeaderPart>(bkmStart, header);
                    }
                }
            }
            return null;
        }

        private BookmarkStart FindBookmarkStartInHeader(HeaderPart header, string name)
        {
            var bkmStart = header.RootElement.Descendants<BookmarkStart>()
                .FirstOrDefault(b => string.Equals(b.Name.ToString(), name, StringComparison.OrdinalIgnoreCase));
            return bkmStart;
        }

        private BookmarkEnd FindBookmarkEndInHeader(HeaderPart header, string idBkm)
        {
            var bkmEnd = header.RootElement.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => string.Equals(b.Id, idBkm, StringComparison.OrdinalIgnoreCase));
            return bkmEnd;
        }

        private Tuple<BookmarkStart, FooterPart> FindBookmarkStartInFooters(string name)
        {
            if (Doc.MainDocumentPart.FooterParts != null)
            {
                foreach (var footer in Doc.MainDocumentPart.FooterParts)
                {
                    var bkmStart = FindBookmarkStartInFooter(footer, name);
                    if (bkmStart != null)
                    {
                        return new Tuple<BookmarkStart, FooterPart>(bkmStart, footer);
                    }
                }
            }
            return null;
        }

        private BookmarkStart FindBookmarkStartInFooter(FooterPart footer, string name)
        {
            var bkmStart = footer.RootElement.Descendants<BookmarkStart>()
                .FirstOrDefault(b => string.Equals(b.Name.ToString(), name, StringComparison.OrdinalIgnoreCase));
            return bkmStart;
        }

        private BookmarkEnd FindBookmarkEndInFooter(FooterPart footer, string idBkm)
        {
            var bkmEnd = footer.RootElement.Descendants<BookmarkEnd>()
                .FirstOrDefault(b => string.Equals(b.Id, idBkm, StringComparison.OrdinalIgnoreCase));
            return bkmEnd;
        }

        private string GetText()
        {
            if (BkmStart != null && BkmEnd != null)
            {
                var text = GetBookmarkText();
                if (text != null)
                    return text.Text;
            }
            return string.Empty;
        }

        private Text GetBookmarkText()
        {
            if (Table != null)
                return GetTextInColumn();
            if (BkmStart.Parent != BkmEnd.Parent)
                return GetMoreParagrathsText();
            return GetOneParagraphText();
        }

        private Text GetTextInColumn()
        {
            var text = new Text();
            var cell = BkmStart.GetParent<TableCell>();
            var res = new StringBuilder();
            var list = cell.Descendants<Text>();
            foreach (var t in list)
            {
                res.AppendLine(t.Text);
            }
            text.Text = res.ToString().TrimEnd();
            return text;
        }

        private Text GetMoreParagrathsText()
        {
            var text = new Text();
            var res = new StringBuilder();
            res.AppendLine(BkmStart.Parent.GetFirstDescendant<Text>().Text);
            var list = GetBookmarkRangeElements();
            foreach (var t in list.Where(el => el.GetType() == typeof(Paragraph)))
            {
                res.AppendLine(t.GetFirstDescendant<Text>().Text);
            }
            text.Text = res.ToString().TrimEnd();
            return text;
        }

        private Text GetOneParagraphText()
        {
            var text = new Text();
            var res = new StringBuilder();
            var list = GetBookmarkElements();
            foreach (var t in list.Where(el => el.GetType() == typeof(Paragraph)))
            {
                res.AppendLine(t.GetFirstDescendant<Text>().Text);
            }
            if (res.Length == 0)
                res.AppendLine(BkmStart.Parent.GetFirstDescendant<Text>().Text);
            text.Text = res.ToString().TrimEnd();
            return text;
        }

        private void SetText(string text)
        {
            if (BkmStart == null || BkmEnd == null)
            {
                return;
            }
            if (BkmEnd.Parent.GetType() == typeof(Table))
            {
                ReplaceTable(text);
                return;
            }
            SiblingElement = BkmStart.Parent.PreviousSibling();
            if (SiblingElement == null)
            {
                SiblingElement = BkmStart.Parent.Parent;
            }
            if (BkmStart.Parent != BkmEnd.Parent)
            {
                ReplaceMoreParagraphs(text);
                return;
            }
            ReplaceOneParagraph(text);
        }

        private void ReplaceTable(string text)
        {
            var parentPara =
                (Paragraph)BkmEnd.Parent.ElementsBefore().LastOrDefault(e => e.GetType() == typeof(Paragraph));
            SiblingElement = parentPara;
            BkmEnd.Parent.Remove();
            Table = null;
            BkmStart = new BookmarkStart
            {
                Name = BkmStart.Name,
                Id = BkmStart.Id
            };
            var newPara = new Paragraph();
            if (parentPara?.Descendants<ParagraphProperties>().Any() != null)
            {
                foreach (var props in parentPara.Descendants<ParagraphProperties>())
                {
                    newPara.AppendChild((ParagraphProperties)props.Clone());
                }
            }
            newPara.Append(BkmStart);
            var nRun = new Run();
            var nText = new Text
            {
                Text = text
            };
            nRun.Append(nText);
            newPara.Append(nRun);
            BkmEnd = new BookmarkEnd
            {
                Id = BkmStart.Id
            };
            newPara.Append(BkmEnd);
            InsertXmlELement(newPara);
        }

        private void ReplaceMoreParagraphs(string text)
        {
            var paraProp = (ParagraphProperties)BkmStart.GetParent<Paragraph>().ParagraphProperties.Clone();
            var list = GetBookmarkRangeElements().ToList();
            for (var n = list.Count; n > 0; n--)
            {
                list[n - 1].Remove();
            }
            BkmStart.Parent.Remove();
            if (IsOneParagraphText(text))
            {
                BkmStart = new BookmarkStart
                {
                    Name = BkmStart.Name,
                    Id = BkmStart.Id
                };
                var newPara = new Paragraph
                {
                    ParagraphProperties = (ParagraphProperties)paraProp.Clone()
                };
                newPara.Append(BkmStart);
                var nRun = new Run();
                var nText = new Text
                {
                    Text = text
                };
                nRun.Append(nText);
                newPara.Append(nRun);
                BkmEnd = new BookmarkEnd
                {
                    Id = BkmStart.Id
                };
                newPara.Append(BkmEnd);
                InsertXmlELement(newPara);
            }
            else
            {
                var rProp = GetDefaultRunProperties();
                InsertMultiLineText(text, SiblingElement, paraProp, rProp);
            }
        }

        private void ReplaceOneParagraph(string text)
        {
            var rProp = GetDefaultRunProperties();
            if (IsOneParagraphText(text))
            {
                if (BkmStart.PreviousSibling<Run>() == null && BkmEnd.ElementsAfter() != null
                    && BkmEnd.ElementsAfter().All(e => e.GetType() != typeof(Run)))
                {
                    BkmStart.Parent.RemoveAllChildren<Run>();
                }
                else
                {
                    var list = GetBookmarkElements().ToList();
                    var trRun = list
                        .Where(rp => rp.GetType() == typeof(Run) && ((Run)rp).RunProperties != null)
                        .Select(rp => ((Run)rp).RunProperties).FirstOrDefault();
                    if (trRun != null)
                    {
                        rProp = (RunProperties)trRun.Clone();
                    }
                    for (var n = list.Count; n > 0; n--)
                    {
                        list[n - 1].Remove();
                    }
                }
                var nRun = new Run();
                if (rProp != null)
                {
                    nRun.RunProperties = (RunProperties)rProp.Clone();
                }
                var nText = new Text
                {
                    Text = text
                };
                nRun.Append(nText);
                BkmStart.InsertAfterSelf(nRun);
            }
            else
            {
                var paraProp = (ParagraphProperties)BkmStart.GetParent<Paragraph>().ParagraphProperties.Clone();
                BkmStart.Parent.Remove();
                InsertMultiLineText(text, SiblingElement, paraProp, rProp);
            }
        }

        private void InsertMultiLineText(string text, OpenXmlElement siblingElement, ParagraphProperties paraProp,
            RunProperties rProp)
        {
            if (!text.Contains("\r\n")) return;
            var insertElement = siblingElement;
            var textLines = text.Split(new[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            var cnt = 0;
            foreach (var textLine in textLines)
            {
                var newPara = new Paragraph
                {
                    ParagraphProperties = (ParagraphProperties)paraProp.Clone()
                };
                if (cnt == 0)
                {
                    BkmStart = new BookmarkStart
                    {
                        Name = BkmStart.Name,
                        Id = BkmStart.Id
                    };
                    newPara.AppendChild(BkmStart);
                }
                var nr = new Run();
                if (rProp != null)
                {
                    nr.RunProperties = (RunProperties)rProp.Clone();
                }
                nr.AppendChild(new Text(textLine));
                newPara.AppendChild(nr);
                if (cnt == textLines.Length - 1)
                {
                    BkmEnd = new BookmarkEnd
                    {
                        Id = BkmStart.Id
                    };
                    newPara.AppendChild(BkmEnd);
                }
                if (insertElement == null)
                {
                    BkmStart.Parent.InsertAfter(newPara, BkmStart);
                }
                else
                {
                    if (insertElement.Parent != null)
                    {
                        insertElement.InsertAfterSelf(newPara);
                    }
                    else
                    {
                        insertElement.Append(newPara);
                    }
                }
                insertElement = newPara;
                cnt += 1;
            }
        }

        private bool IsOneParagraphText(string text)
        {
            return string.IsNullOrEmpty(text) || (!text.Contains("\r") && !text.Contains("\n"));
        }

        private RunProperties GetDefaultRunProperties()
        {
            RunProperties rProp = null;
            var r = BkmStart.GetParent<Run>();
            if (r != null)
            {
                rProp = r.GetFirstDescendant<RunProperties>();
            }
            if (rProp != null)
            {
                return rProp;
            }
            var p = BkmStart.GetParent<Paragraph>();
            var pProp = p?.GetFirstDescendant<ParagraphMarkRunProperties>();
            if (pProp == null)
            {
                return null;
            }
            rProp = new RunProperties();
            foreach (var pr in pProp)
            {
                rProp.AppendChild(pr.CloneNode(true));
            }
            return rProp;
        }

        private void InsertXmlELement(OpenXmlElement element)
        {
            if (SiblingElement != null)
            {
                SiblingElement.InsertAfterSelf(element);
            }
            else
            {
                BkmStart.GetParent<Paragraph>().InsertAfter(element, BkmStart);
            }

        }

        private IEnumerable<OpenXmlElement> GetBookmarkRangeElements()
        {
            var list = BkmStart.Parent.ElementsAfter()
                .Where(p => p.IsBefore(BkmEnd.Parent) || p == BkmEnd.Parent);
            return list;
        }

        private IEnumerable<OpenXmlElement> GetBookmarkElements()
        {
            var list = BkmStart.ElementsAfter()
                .Where(p => p.IsBefore(BkmEnd));
            return list;
        }

        private void TryFindTable()
        {
            if (BkmStart == null || BkmEnd == null)
            {
                return;
            }
            var table = BkmEnd.GetParent<Table>();
            if (table != null)
            {
                Table = new WordDocumentTable(Doc, table);
            }
        }

        private Tuple<BookmarkStart, BookmarkEnd> GetBookmarkCopy()
        {
            if (BkmStart == null || BkmEnd == null)
            {
                return null;
            }
            int id;
            if (!int.TryParse(BkmStart.Id, out id))
            {
                return null;
            }
            var newBkmStart = (BookmarkStart)BkmStart.Clone();
            newBkmStart.Name = BkmStart.Name + "_copy";
            newBkmStart.Id = (id + 1000).ToString();
            var newBkmEnd = (BookmarkEnd)BkmEnd.Clone();
            newBkmEnd.Id = newBkmStart.Id;
            return new Tuple<BookmarkStart, BookmarkEnd>(newBkmStart, newBkmEnd);
        }

        private bool TryCopyTable(IPrintObject sourcePrintObject)
        {
            try
            {
                var tbl = sourcePrintObject as IWordDocumentTable;
                if (tbl != null)
                {
                    CopyTableToBookmark(tbl);
                    return true;
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void CopyTableToBookmark(IWordDocumentTable sourceTable)
        {
            var parentPara =
                (Paragraph)BkmEnd.Parent.ElementsBefore().LastOrDefault(e => e.GetType() == typeof(Paragraph));
            SiblingElement = parentPara;
            BkmEnd.Parent.Remove();
            Table = null;
            BkmStart = new BookmarkStart
            {
                Name = BkmStart.Name,
                Id = BkmStart.Id
            };
            BkmEnd = new BookmarkEnd
            {
                Id = BkmStart.Id
            };
            var newTable = (Table) sourceTable.GetCopy().Table;
            InsertXmlELement(newTable);
            newTable.GetFirstChild<TableRow>().GetFirstChild<TableCell>().Append(BkmStart);
            newTable.AppendChild(BkmEnd);
            Table = new WordDocumentTable(Doc, newTable);
        }
    }
}