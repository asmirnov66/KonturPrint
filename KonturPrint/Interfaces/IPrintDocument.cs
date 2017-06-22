using SKBS;

namespace KonturPrint.Interfaces
{
    public interface IPrintDocument
    {
        bool IsSameDocumentType(PrintDocumentType type);
        void LoadTemplate(string filePath);
        void LoadTemplate(byte[] fileContent);
        void AddBo(string boName, IBSDataObject bo);
        void SetActiveBo(string boName);
        bool CheckForInitialization();
        byte[] Print();
        string PrintToPath(string path);
        void ProcessDocument();
        bool Update();
    }
}