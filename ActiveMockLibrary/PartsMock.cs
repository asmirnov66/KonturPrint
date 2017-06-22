using System;
using System.Collections.Generic;
using SKBS;

namespace ActiveMockLibrary
{
    public class PartsMock : SKBS.Parts
    {
        private Dictionary<string, SKRecordsetMock> items;

        public PartsMock()
        {
            items = new Dictionary<string, SKRecordsetMock>(StringComparer.OrdinalIgnoreCase);
        }

        public SKBS.SKRecordset GetData(object Index, bool AsyncGetData)
        {
            throw new System.NotImplementedException();
        }

        public int Count { get; }
        public object Names { get; }
        public IBSDataObject Parent { get; }
        public bool get_IsDataGet(object Index)
        {
            throw new System.NotImplementedException();
        }

        public bool get_IsDataGetting(object Index)
        {
            throw new System.NotImplementedException();
        }

        public void AddItem(string key, SKRecordsetMock rs)
        {
            items.Add(key, rs);
        }

        public SKBS.SKRecordset Item(object Index)
        {
            return items[(string)Index];
        }
    }
}
