using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using KonturPrint.Factories;
using KonturPrint.Interfaces;
using KonturPrint.PrintDocuments;
using Ninject;
using Ninject.Modules;
using NLog;
using SKBSInt;
using XMachine;

namespace KonturPrintService
{
    [ComVisible(false)]
    public class KonturPrintServiceModule : NinjectModule
    {
        private IXMachine XMachine { get; }

        public KonturPrintServiceModule()
        {
        }

        public KonturPrintServiceModule(IXMachine xMachine)
        {
            XMachine = xMachine;
        }

        public override void Load()
        {
            Bind<IPrintDocument>().To<WordTemplateDocument>();
            Bind<IPrintDocument>().To<ExcelTemplateDocument>();
            Bind<IPrintDocumentFactory>().To<PrintDocumentFactory>().InSingletonScope();
            Bind<ILogger>().ToConstant(LogManager.GetCurrentClassLogger());

            if (XMachine != null)
            {
                Bind<ISKServiceInit, IPrintService>().To<PrintService>().WithConstructorArgument("xMachine", XMachine);
            }
            else
            {
                Bind<ISKServiceInit, IPrintService>().To<PrintService>();
            }
        }
    }
}