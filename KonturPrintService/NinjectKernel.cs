using Ninject;

namespace KonturPrintService
{
    public static class NinjectKernel
    {
        public static readonly IKernel Instance = new StandardKernel();

        public static void Bind()
        {
            Instance.Load(new KonturPrintServiceModule());
        }
    }
}