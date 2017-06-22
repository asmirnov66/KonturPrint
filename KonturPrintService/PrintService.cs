using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using KonturPrint.Interfaces;
using Ninject;
using NLog;
using NLog.Targets;
using SKBS;
using SKBSInt;
using SKGENERALLib;
using XMachine;

namespace KonturPrintService
{
    [Guid("DEE9FA0D-3F46-4228-9BF4-8B230BC2ED5A")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("KonturPrintService.PrintService")]
    [ComVisible(true)]
    public class PrintService : ISKServiceInit, IPrintService
    {
        private IPrintDocumentFactory Factory { get; }
        private IPrintDocument Document { get; set; }
        private IBServerInternal BsInt { get; set; }
        private List<string> PrintScriptsList { get; set; }

        [Inject]
        public ILogger Logger { get; set; }
        public IXMachine XMachine { get; set; }
        public object PrintScripts { get; set; }

        public PrintService()
        {
            NinjectKernel.Bind();
            Factory = NinjectKernel.Instance.Get<IPrintDocumentFactory>();
            Logger = NinjectKernel.Instance.Get<ILogger>();
            PrintScriptsList = new List<string>();
        }

        public PrintService(IXMachine xMachine)
        {
            NinjectKernel.Bind();
            Factory = NinjectKernel.Instance.Get<IPrintDocumentFactory>();
            XMachine = xMachine;
            Logger = NinjectKernel.Instance.Get<ILogger>();
            PrintScriptsList = new List<string>();
        }

        public PrintService(IXMachine xMachine, ILogger logger)
        {
            NinjectKernel.Bind();
            Factory = NinjectKernel.Instance.Get<IPrintDocumentFactory>();
            XMachine = xMachine;
            Logger = logger;
            PrintScriptsList = new List<string>();
        }

        public PrintService(IPrintDocumentFactory factory)
        {
            Factory = factory;
            PrintScriptsList = new List<string>();
        }

        public PrintService(IPrintDocumentFactory factory, IXMachine xMachine)
        {
            Factory = factory;
            XMachine = xMachine;
            PrintScriptsList = new List<string>();
        }

        public PrintService(IPrintDocumentFactory factory, IXMachine xMachine, ILogger logger)
        {
            Factory = factory;
            XMachine = xMachine;
            Logger = logger;
            PrintScriptsList = new List<string>();
        }

        public void Init(string userName, string configName, object serviceParams = null)
        {
            var p = (IParams)serviceParams;
            BsInt = p?.GetValue("BServerInternal");
            if (BsInt != null)
            {
                InitLogger();
                var useXMachine = p?.GetValue("UseXMachine", true);
                if (useXMachine)
                    XMachine = (IXMachine)BsInt.GetCustomService("XMachine");
                PrintScriptsList = new List<string>();
            }
        }

        public void Close()
        {
            XMachine = null;
        }

        public bool Print(int printDocumentType, object printParams = null)
        {
            PrintDocumentType? docType;
            if (!TryGetPrintDocumentType(printDocumentType, out docType))
            {
                return false;
            }
            var p = (IParams)printParams;
            switch (docType)
            {
                case PrintDocumentType.ExcelTemplate:
                    Document = Factory.GetPrintDocument((PrintDocumentType)docType);
                    break;
                case PrintDocumentType.WordTemplate:
                    Document = Factory.GetPrintDocument((PrintDocumentType)docType);
                    return PrintWordTemplate(p);
                default:
                    return false;
            }
            return true;
        }

        public object GetPrintDocument(int printDocumentType, object printParams = null)
        {
            PrintDocumentType? docType;
            if (!TryGetPrintDocumentType(printDocumentType, out docType))
            {
                return null;
            }
            var p = (IParams)printParams;

            switch (docType)
            {
                case PrintDocumentType.ExcelTemplate:
                    Document = Factory.GetPrintDocument((PrintDocumentType)docType);
                    return Document;
                case PrintDocumentType.WordTemplate:
                    Document = Factory.GetPrintDocument((PrintDocumentType)docType);
                    if (!ValidatePrintParams(p))
                    {
                        return null;
                    }
                    InitWordDocument(p);
                    return (IWordDocument)Document;
                default:
                    return null;
            }
        }

        public void SavePrintDocument(object printParams = null)
        {
            InternalSavePrintDocument((IParams)printParams);
        }

        private bool PrintWordTemplate(IParams p)
        {
            if (!ValidatePrintParams(p))
            {
                return false;
            }
            try
            {
                InitWordDocument(p);
                var printParams = (IParams)p.GetValue("Params");
                FillPrintScriptsList(printParams);
                EvaluatePrintScripts(printParams);
                InternalSavePrintDocument(p);
            }
            catch (Exception e)
            {
                Logger.Error(e.Message + "\r\n" + e.StackTrace);
                p.SetParams("ErrorStr", SaveErrorToString(e));
                return false;
            }
            return true;
        }

        private bool TryGetPrintDocumentType(int docType, out PrintDocumentType? printDocumentType)
        {
            if (Enum.IsDefined(typeof(PrintDocumentType), docType))
            {
                printDocumentType = (PrintDocumentType)docType;
                return true;
            }
            printDocumentType = null;
            return false;
        }

        private bool ValidatePrintParams(IParams p)
        {
            string errStr;
            var templatePath = p.GetValue("TemplatePath", "");
            if (string.IsNullOrEmpty(templatePath))
            {
                errStr = "Не задан файл шаблона";
                p.SetParams("ErrorStr", SaveErrorToString(errStr));
                Logger.Error(errStr);
                return false;
            }

            var fileName = p.GetValue("FileName", "");
            if (string.IsNullOrEmpty(fileName))
            {
                errStr = "Не задан выходной файл";
                p.SetParams("ErrorStr", SaveErrorToString(errStr));
                Logger.Error(errStr);
                return false;
            }

            var bo = p.GetValue("BO");
            if (bo == null)
            {
                errStr = "Не задан бизнес - объект";
                p.SetParams("ErrorStr", SaveErrorToString(errStr));
                Logger.Error(errStr);
                return false;
            }
            return true;
        }

        private void InitLogger()
        {
            var installPath = BsInt.BusinessServer.RunCommand("GetInstallPath");
            var logfile = installPath + @"\Logs\KPS.log";
            var target = LogManager.Configuration.FindTargetByName("logfile") as FileTarget;
            if (target != null) target.FileName = logfile;
        }

        private string SaveErrorToString(string error)
        {
            return "-2147024809" + (char)1 + "PrintService" + (char)1 + error + (char)1 + "c:\\" + (char)1 + "0";
        }

        private string SaveErrorToString(Exception error)
        {
            return error.HResult.ToString() + (char)1 + error.Source + (char)1 + error.Message + (char)1 + error.HelpLink + (char)1 + "0";
        }

        private void EvaluatePrintScripts(IParams printParams)
        {
            foreach (var block in PrintScriptsList)
            {
                EvaluatePrintScript(block, printParams);
            }
        }

        private void EvaluatePrintScript(string block, IParams printParams)
        {
            if (!string.IsNullOrEmpty(block))
                XMachine.Evaluate(block, "ActiveDocument", (IWordDocument)Document, "printParams", printParams);
        }

        private void FillPrintScriptsList(IParams printParams)
        {
            PrintScriptsList = new List<string>();
            var block = printParams.GetValue("PrintScript", "");
            if (!string.IsNullOrEmpty(block))
            {
                PrintScriptsList.Add(block);
            }
            var obj = PrintScripts as object[];
            if (obj == null)
            {
                return;
            }
            PrintScriptsList.AddRange(obj.Cast<string>());
        }

        private void InitWordDocument(IParams p)
        {
            var templatePath = p.GetValue("TemplatePath", "");
            var bo = (IBSDataObject)p.GetValue("BO");
            Document.AddBo(bo.Name, bo);
            Document.LoadTemplate(templatePath);
            Document.ProcessDocument();
        }

        private void InternalSavePrintDocument(IParams p)
        {
            var fileName = p.GetValue("FileName", "");
            if (Document.Update())
            {
                Document.PrintToPath(fileName);
            }
        }
    }
}
