using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects.PageEnumerators
{
    public class WordDocumentPageEnumerator : IPrintObject
    {
        public OpenXmlElement XmlElement { get; }

        public WordDocumentPageEnumerator()
        {
            XmlElement = GetXmlElement();
        }

        public IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            var sourceObject = new PrintObject(XmlElement);
            return sourceObject.CopyTo(destPrintObject);
        }

        public IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            return null;
        }

        private OpenXmlElement GetXmlElement()
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties(
                new ParagraphStyleId
                {
                    Val = "Footer"
                },
                new Justification
                {
                    Val = JustificationValues.Right
                }
            );
            paragraph.Append(paragraphProperties);
            paragraph.Append(new Run(
                new Text
                {
                    Text = "стр. ",
                    Space = SpaceProcessingModeValues.Preserve
                }
            ));
            paragraph.Append(new Run(
                new FieldChar
                {
                    FieldCharType = FieldCharValues.Begin
                }
            ));
            paragraph.Append(new Run(
                new FieldCode
                {
                    Space = SpaceProcessingModeValues.Preserve,
                    Text = "PAGE"
                },
                new RunProperties(new RunStyle
                {
                    Val = "PageNumber"
                }
            )));
            paragraph.Append(new Run
                (new FieldChar
                {
                    FieldCharType = FieldCharValues.End
                }
            ));
            paragraph.Append(new Run(
                new Text
                {
                    Text = " из ",
                    Space = SpaceProcessingModeValues.Preserve
                }
            ));
            paragraph.Append(new Run(
                new FieldChar
                {
                    FieldCharType = FieldCharValues.Begin
                }
            ));
            paragraph.Append(new Run(
                new FieldCode
                {
                    Space = SpaceProcessingModeValues.Preserve,
                    Text = "NUMPAGES"
                },
                new RunProperties(new RunStyle
                {
                    Val = "PageNumber"
                }
            )));
            paragraph.Append(new Run(
                new FieldChar
                {
                    FieldCharType = FieldCharValues.End
                }
            ));
            var pageNumberBlock = new SdtBlock();
            var contentBlock = new SdtContentBlock();
            contentBlock.Append(paragraph);
            pageNumberBlock.Append(contentBlock);
            return pageNumberBlock;
        }
    }
}