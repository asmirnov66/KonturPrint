using System;
using SKBS;

namespace ActiveMockLibrary
{
    public class SKFieldMock : SKBS.SKField
    {
        public void let_Value(object Result)
        {
            throw new System.NotImplementedException();
        }

        public object GetProperty(string Name, object DefaultValue)
        {
            throw new System.NotImplementedException();
        }

        public void SetValueEx(ref object Value, object Options)
        {
            throw new System.NotImplementedException();
        }

        public object GetEditService(SKFieldEditServiceEnum EditService)
        {
            throw new System.NotImplementedException();
        }

        public void SetTextEx(string Value, ref object Options)
        {
            throw new System.NotImplementedException();
        }

        public object Value { get; set; }
        public string Name { get; }
        public string Caption { get; }
        public bool Visible { get; }
        public bool Required { get; }
        public int ActualSize { get; }
        public int DefinedSize { get; }
        public object OriginalValue { get; }
        public object OldValue { get; }

        public int Type
        {
            get
            {
                var outInt = 0;
                if (int.TryParse(DisplayText, out outInt))
                    return 3;
                return 200;
            }
        }

        public string DisplayText { get { return Convert.ToString(Value); } }
        public string Text { get; set; }
        public string Reference { get; }
        public bool IsChangeInTran { get; }
        public int Index { get; }
        public SKFields Parent { get; }
        public bool IsChange { get; }
    }

}
