using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using KonturPrint.Interfaces;
using SKBS;

namespace KonturPrint.PrintDocuments
{
    public abstract class BaseDocument : IPrintDocument
    {
        protected Dictionary<string, IBSDataObject> BoDictionary;
        protected string ActiveBoName;
        protected IBSDataObject ActiveBo;
        protected MemoryStream MemStream;
        protected OpenXmlPackage Doc;

        protected BaseDocument()
        {
            BoDictionary = new Dictionary<string, IBSDataObject>(StringComparer.OrdinalIgnoreCase);
        }

        ~BaseDocument()
        {
            Doc?.Dispose();
            MemStream?.Dispose();
        }

        public virtual void LoadTemplate(string filePath)
        {
            var byteArray = File.ReadAllBytes(filePath);
            MemStream = new MemoryStream();
            MemStream.Write(byteArray, 0, byteArray.Length);
        }

        public virtual void LoadTemplate(byte[] fileContent)
        {
            MemStream = new MemoryStream(fileContent);
        }

        public virtual void AddBo(string boName, IBSDataObject bo)
        {
            BoDictionary.Add(boName, bo);
            if (ActiveBoName == null)
            {
                SetActiveBo(boName);
            }
        }

        public virtual byte[] Print()
        {
            return MemStream.ToArray();
        }

        public virtual string PrintToPath(string path)
        {
            using (var fileStream = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite))
            {
                MemStream.WriteTo(fileStream);
            }
            return path;
        }

        public abstract void ProcessDocument();

        public virtual bool CheckForInitialization()
        {
            if (MemStream == null)
                throw new Exception("Не произведена инициализация документа");
            if (ActiveBo == null)
                throw new Exception("Не задан бизнес - объект");
            return true;
        }

        public virtual void SetActiveBo(string boName)
        {
            ActiveBo = BoDictionary[boName];
            ActiveBoName = boName;
        }

        public abstract bool Update();
        public abstract bool IsSameDocumentType(PrintDocumentType type);
    }
}
