using SKBS;
using SKGENERALLib;

namespace ActiveMockLibrary
{
    public class BsDataObjectMock : SKBS.IBSDataObject
    {
        private SKBS.Parts parts;

        public string Name { get; }

        public BsDataObjectMock()
        {
            parts = new PartsMock();
        }

        public BsDataObjectMock(string name)
        {
            Name = name;
            parts = new PartsMock();
        }
        object SKBS.IBSDataObject.BOInternal { get; }
        SKBS.Parts SKBS.IBSDataObject.Parts { get { return parts; } }
        public bool IsChange { get; set; }
        public bool ReadOnly { get; }

        void IBSObject.CloseObject()
        {
            throw new System.NotImplementedException();
        }

        object SKBS.IBSDataObject.RunCommand(string Name, object Params)
        {
            throw new System.NotImplementedException();
        }

        Params SKBS.IBSDataObject.GetCommandParams(string CmdName)
        {
            throw new System.NotImplementedException();
        }

        object SKBS.IBSDataObject.GetPropertyValue(string PropName, object DefaultValue)
        {
            throw new System.NotImplementedException();
        }

        void SKBS.IBSDataObject.Init(object BServerInternal, string ObjName, object Node, Params Params, Params Options)
        {
            throw new System.NotImplementedException();
        }

        void SKBS.IBSDataObject.ReOpen(ref object Params, ref object Options)
        {
            throw new System.NotImplementedException();
        }

        bool SKBS.IBSDataObject.TryRunCommand(string Name, object Params, ref object Result)
        {
            throw new System.NotImplementedException();
        }

        int SKBS.IBSDataObject.ObjectType { get; }
        string SKBS.IBSDataObject.Caption { get; }
        BusinessServer SKBS.IBSDataObject.BusinessServer { get; }

        public void Update()
        {
            throw new System.NotImplementedException();
        }

        void SKBS.IBSDataObject.CloseObject()
        {
            throw new System.NotImplementedException();
        }

        object IBSObject.RunCommand(string Name, object Params)
        {
            throw new System.NotImplementedException();
        }

        Params IBSObject.GetCommandParams(string CmdName)
        {
            throw new System.NotImplementedException();
        }

        object IBSObject.GetPropertyValue(string PropName, object DefaultValue)
        {
            throw new System.NotImplementedException();
        }

        void IBSObject.Init(object BServerInternal, string ObjName, object Node, Params Params, Params Options)
        {
            throw new System.NotImplementedException();
        }

        void IBSObject.ReOpen(ref object Params, ref object Options)
        {
            throw new System.NotImplementedException();
        }

        bool IBSObject.TryRunCommand(string Name, object Params, ref object Result)
        {
            throw new System.NotImplementedException();
        }

        string IBSObject.Name { get; }
        int IBSObject.ObjectType { get; }
        string IBSObject.Caption { get; }
        BusinessServer IBSObject.BusinessServer { get; }
        bool IBSObject.get_Permissions(string Permission)
        {
            throw new System.NotImplementedException();
        }

        Params SKBS.IBSDataObject.Params { get; }

        bool SKBS.IBSDataObject.get_Permissions(string Permission)
        {
            throw new System.NotImplementedException();
        }

        Params IBSObject.Params { get; }
        object IBSObject.BOInternal { get; }
    }
}
