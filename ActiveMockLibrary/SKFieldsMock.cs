using System;
using System.Collections;
using System.Collections.Specialized;
using SKBS;

namespace ActiveMockLibrary
{
    public class SKFieldsMock : SKBS.SKFields
    {
        private OrderedDictionary fields;

        public SKFieldsMock()
        {
            fields = new OrderedDictionary();
        }

        public void AddField(object key, object value)
        {
            if (!fields.Contains(key))
            {
                fields.Add(key, new SKFieldMock { Value = value });
            }
        }

        public void Append(string Name, int DataType)
        {
            throw new NotImplementedException();
        }

        public void Delete(string Name)
        {
            throw new NotImplementedException();
        }

        public int getFieldIndex(string FieldName)
        {
            throw new NotImplementedException();
        }

        IEnumerator SKFields.GetEnumerator()
        {
            throw new NotImplementedException();
        }

        public SKField this[object index]
        {
            get
            {
                var t = index.GetType();
                if (t == typeof(string))
                {
                    return (SKField)fields[index];
                }
                var size = fields.Keys.Count;
                var values = new SKFieldMock[size];
                fields.Values.CopyTo(values, 0);
                return values[int.Parse(index.ToString())];
            }
        }

        public int Count
        {
            get
            {
                return fields.Count;
            }
        }

        public object Names { get; }
        public SKRecordset Parent { get; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            throw new NotImplementedException();
        }
    }
}
