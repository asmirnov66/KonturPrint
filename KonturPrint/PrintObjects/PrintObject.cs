using DocumentFormat.OpenXml;
using KonturPrint.Interfaces;

namespace KonturPrint.PrintObjects
{
    public class PrintObject : IPrintObject
    {
        public OpenXmlElement XmlElement { get; }

        public PrintObject(OpenXmlElement xmlElement)
        {
            XmlElement = xmlElement;
        }

        public PrintObject(object xmlElement)
        {
            XmlElement = (OpenXmlElement)xmlElement;
        }

        public IPrintObject CopyTo(IPrintObject destPrintObject)
        {
            if (XmlElement == null)
            {
                return null;
            }
            var sourceObject = (OpenXmlElement)XmlElement.Clone();
            var destObject = destPrintObject.XmlElement.AppendChild(sourceObject);
            return new PrintObject(destObject);
        }

        public IPrintObject GetCopyOf(IPrintObject sourcePrintObject)
        {
            if (sourcePrintObject.XmlElement == null)
            {
                return null;
            }
            var sourceObject = (OpenXmlElement)sourcePrintObject.XmlElement.Clone();
            var destObject = XmlElement.AppendChild(sourceObject);
            return new PrintObject(destObject);
        }
    }
}