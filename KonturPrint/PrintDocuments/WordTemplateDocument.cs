using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KonturPrint.Interfaces;
using KonturPrint.PrintObjects.Bookmarks;
using KonturPrint.PrintObjects.HeadersFooters.Footers;
using KonturPrint.PrintObjects.HeadersFooters.Headers;
using KonturPrint.PrintObjects.PageEnumerators;
using KonturPrint.PrintObjects.Tables;

namespace KonturPrint.PrintDocuments
{
    public class WordTemplateDocument : BaseDocument, IWordDocument, IWordDocumentStructure
    {
        public IWordDocumentBookmarks Bookmarks { get; private set; }
        public IWordDocumentTables Tables { get; private set; }
        public IWordDocumentHeadersFooters Footers { get; private set; }
        public IWordDocumentHeadersFooters Headers { get; private set; }
        public IPrintObject PageEnumerator { get; private set; }

        public WordprocessingDocument InnerDoc { get; private set; }
        public SectionProperties SectionProperties
        {
            get { return GetSectionProperties(); }
        }
        public Settings Settings
        {
            get { return GetSettings(); }
        }

        public WordTemplateDocument()
        {
        }

        public WordTemplateDocument(WordprocessingDocument doc)
        {
            InnerDoc = doc;
        }

        public override bool IsSameDocumentType(PrintDocumentType type)
        {
            return type == PrintDocumentType.WordTemplate;
        }

        public override void ProcessDocument()
        {
            CheckForInitialization();
            Doc = WordprocessingDocument.Open(MemStream, true);
            InnerDoc = (WordprocessingDocument)Doc;
            Bookmarks = new WordDocumentBookmarks(this);
            Tables = new WordDocumentTables(this);
            Footers = new WordDocumentFooters(this);
            Headers = new WordDocumentHeaders(this);
            PageEnumerator = new WordDocumentPageEnumerator();
        }

        public override bool Update()
        {
            if (Doc == null)
                return false;
            var wordDoc = (WordprocessingDocument)Doc;
            var mainPart = wordDoc.MainDocumentPart;
            mainPart.Document.Save();
            foreach (var f in mainPart.FooterParts)
            {
                f.Footer.Save();
            }
            foreach (var h in mainPart.HeaderParts)
            {
                h.Header.Save();
            }
            Doc.Close();
            return true;
        }


        private SectionProperties GetSectionProperties()
        {
            var body = InnerDoc.MainDocumentPart.Document.Body;
            var sectionProperties = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectionProperties == null)
            {
                sectionProperties = new SectionProperties();
                body.Append(sectionProperties);
            }
            return sectionProperties;
        }

        private Settings GetSettings()
        {
            var mainDocumentPart = InnerDoc.MainDocumentPart;
            var settingsPart = mainDocumentPart.DocumentSettingsPart;
            if (settingsPart == null)
            {
                settingsPart = mainDocumentPart.AddNewPart<DocumentSettingsPart>();
            }
            return settingsPart.Settings;
        }
    }
}