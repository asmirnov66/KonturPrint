using System;
using KonturPrint.Factories;
using KonturPrint.Interfaces;
using KonturPrint.PrintDocuments;
using Ninject;
using NLog;

namespace KonturPrintTest
{
    public static class NinjectKernel
    {
        public static readonly IKernel Instance = new StandardKernel();

        public static void Bind()
        {
            Instance.Rebind<IPrintDocument>().To<WordTemplateDocument>();
            Instance.Rebind<IPrintDocument>().To<ExcelTemplateDocument>();
            Instance.Rebind<IPrintDocumentFactory>().To<PrintDocumentFactory>().InSingletonScope();
            Instance.Rebind<ILogger>().ToConstant(LogManager.GetCurrentClassLogger());
        }

        public static void UnBind(Type service)
        {
            Instance.Unbind(service);
        }
    }
}