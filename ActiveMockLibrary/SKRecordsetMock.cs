using System;
using System.Collections.Generic;
using System.Linq;
using SKBS;

namespace ActiveMockLibrary
{
    public class SKRecordsetMock : SKBS.SKRecordset
    {
        private int bookMark = -1;
        private List<SKFields> fields;
        private List<SKFields> filteredFields;
        private string filterString;

        public SKRecordsetMock()
        {
            fields = new List<SKFields>();
            filteredFields = new List<SKFields>();
        }

        public SKBS.SKFields Fields
        {
            get
            {
                if (bookMark == -1)
                    throw new IndexOutOfRangeException("Out of range Exception");
                return filteredFields.ElementAt(bookMark);
            }
        }
        public string Name { get; }
        public int Index { get; }
        public int Bookmark { get { return bookMark; } set { bookMark = value; } }
        public string Sort { get; set; }
        public bool Filtered { get; }
        public bool EOF
        {
            get
            {
                return bookMark > filteredFields.Count - 1;
            }
        }
        public bool BOF
        {
            get
            {
                return bookMark < 0;
            }
        }
        public bool IsChange { get; }
        public bool ReadOnly { get; }
        public bool IsDataGet { get; }
        public bool IsDataActual { get; }
        public int AbsolutePosition { get; set; }
        public int RecordCount { get; }
        public IBSDataObject BSDataObject { get; }
        public object DataSource { get; }
        public object AddedRows { get; }
        public object EditedRows { get; }
        public object DeletedRows { get; }
        public int RowStatus { get; }
        public SKRsStateEnum State { get; }
        public Selection Selection { get; }
        public Filters Filters { get; }
        public object AddedRowsInTran { get; }
        public object EditedRowsInTran { get; }
        public object DeletedRowsInTran { get; }
        public int RowStatusInTran { get; }

        public void AddNew(object Fields, object Values)
        {
            var k = Fields as object[];
            var v = Values as object[];
            if (k == null || k.Length == 0)
                throw new Exception("Нельзя добавить пустой кортеж");
            if (v == null || k.Length != v.Length)
                throw new Exception("Не совпадает количество полей и значений");
            var skFields = new SKFieldsMock();
            for (var i = 0; i < k.Length; i++)
            {
                skFields.AddField(k[i], v[i]);
            }
            fields.Add(skFields);
            SetFilter(filterString, false);
            bookMark = filteredFields.Count - 1;
        }
        public object GetPropertyBatch(string Name, ref object Bookmarks, ref object Fields)
        {
            throw new NotImplementedException();
        }

        public void DeleteRecord(bool AllRecords)
        {
            throw new NotImplementedException();
        }

        public void Move(int NumRecords, int StartBook = 0, int StartPos = 0)
        {
            throw new NotImplementedException();
        }

        public void MoveFirst()
        {
            if (bookMark == -1)
                throw new Exception("Нет данных!");
            bookMark = 0;
        }

        public void MoveLast()
        {
            if (bookMark == -1 || filteredFields.Count == 0)
                throw new Exception("Нет данных!");
            bookMark = filteredFields.Count - 1;
        }

        public void MoveNext()
        {
            if (BOF)
                throw new Exception("Нет данных!");
            if (EOF)
                throw new Exception("Достигнут конец Recordset'a");
            bookMark++;
        }

        public void MovePrevious()
        {
            throw new NotImplementedException();
        }

        public void Update()
        {
            throw new NotImplementedException();
        }

        public void Requery(bool Force, bool AsyncGetData)
        {
            throw new NotImplementedException();
        }

        public void SetFieldValues(object Fields, object Values)
        {
            throw new NotImplementedException();
        }

        public SKRecordset Clone(bool ReadOnly, bool CopyFilerAndSort)
        {
            throw new NotImplementedException();
        }

        public void Open(object Source, object Options)
        {
            throw new NotImplementedException();
        }

        public void SetFilter(object FilterExp, bool OnExist)
        {
            var fStr = (string)FilterExp;
            filteredFields.Clear();
            if (string.IsNullOrEmpty(fStr))
            {
                foreach (var row in fields)
                    filteredFields.Add(row);
                return;
            }
            var values = fStr.Split('=');
            if (values.Length == 0)
            {
                throw new NotImplementedException("Поддерживается только фильтр '='");
            }
            var fieldName = values[0];
            var fieldValue = values[1];

            foreach (var row in fields)
            {
                if (row[fieldName].DisplayText == fieldValue)
                    filteredFields.Add(row);
            }
            bookMark = filteredFields.Count - 1;
        }

        public void ValidateRecord()
        {
            throw new NotImplementedException();
        }

        public bool IsRecordExists(int Bookmark)
        {
            throw new NotImplementedException();
        }

        public int CompareBookmarks(int Bookmark1, int Bookmark2)
        {
            throw new NotImplementedException();
        }

        public object GetProperty(string Name, object DefaultValue)
        {
            throw new NotImplementedException();
        }

        public int Int1(int p1, int p2, int p3)
        {
            throw new NotImplementedException();
        }

        public void CancelGetData()
        {
            throw new NotImplementedException();
        }

        public object CalcAggregates(object Fields, object AggregatesKind, ref object Bookmarks)
        {
            throw new NotImplementedException();
        }

        public object Sum(object Fields)
        {
            throw new NotImplementedException();
        }

        public object GetBookmarks(int StartBook = -1, int StartPos = 0, bool Forward = true, int Count = 0)
        {
            throw new NotImplementedException();
        }

        public bool Find(string Criteria, bool Forward = true, int StartBook = 0, int StartPos = 0)
        {
            throw new NotImplementedException();
        }

        public void Delete(ref object Bookmarks)
        {
            throw new NotImplementedException();
        }

        public SKRecordset CloneEx(SKCloneOptionsEnum Options = (SKCloneOptionsEnum)0)
        {
            throw new NotImplementedException();
        }

        public int GetBookmarkPosition(int Bookmark)
        {
            throw new NotImplementedException();
        }

        public dynamic GetFieldValues(object Fields, int Records, bool NoChangePos, string Property, ref object StartBook, ref object StartPos)
        {
            throw new NotImplementedException();
        }

        public dynamic GetRows(object Fields, int Records, bool NoChangePos, string Property, ref object StartBook, ref object StartPos)
        {
            throw new NotImplementedException();
        }

        public dynamic GetTargetRows(object Fields, bool NoChangePos, string Property, ref object TargetRecords, ref object CollectedBookmarks)
        {
            throw new NotImplementedException();
        }

        public event ISKRecordsetEvents_DataChangedEventHandler DataChanged;
        public event ISKRecordsetEvents_MoveCompleteEventHandler MoveComplete;
        public event ISKRecordsetEvents_DataGetEventHandler DataGet;
        public event ISKRecordsetEvents_ErrorGetDataEventHandler ErrorGetData;
        public event ISKRecordsetEvents_SelectionChangedEventHandler SelectionChanged;
        public event ISKRecordsetEvents_BeforeRequeryEventHandler BeforeRequery;
    }
}
