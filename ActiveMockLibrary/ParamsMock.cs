using SKGENERALLib;

namespace ActiveMockLibrary
{
    public class ParamsMock
    {
        private IParams Params { get; }

        public ParamsMock(IParams p)
        {
            Params = p;
        }

        public ParamsMock(object p)
        {
            Params = (IParams)p;
        }

        public dynamic GetValue(object paramIndex, object defaultValue)
        {
            return Params.GetValue(paramIndex, defaultValue);
        }
    }
}