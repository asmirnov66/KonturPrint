using System.Runtime.InteropServices;

namespace KonturPrint.Interfaces
{
    public enum HeadersFootersType
    {
        EvenOdd,
        Default,
        FirstDefault,
        FirstEvenOdd
    }

    [Guid("2B760DD7-6E20-48F3-849F-7B49150DB5E8")]
    [ComVisible(true)]
    public interface IWordDocumentHeadersFooters
    {
        int Count { get; }

        IWordDocumentHeaderFooter Item(int index);
        IWordDocumentHeaderFooter ItemByType(int type);
        void Create(int type);
    }
}