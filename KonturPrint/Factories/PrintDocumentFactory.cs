using System.Collections.Generic;
using KonturPrint.Interfaces;

namespace KonturPrint.Factories
{
    public class PrintDocumentFactory : IPrintDocumentFactory
    {
        private readonly IEnumerable<IPrintDocument> printDocuments;

        public PrintDocumentFactory(IEnumerable<IPrintDocument> ptDocuments)
        {
            printDocuments = ptDocuments;
        }

        public IPrintDocument GetPrintDocument(PrintDocumentType documentType)
        {
            foreach (var factory in printDocuments)
            {
                if (factory.IsSameDocumentType(documentType))
                    return factory;
            }
            return null;
        }
    }
}