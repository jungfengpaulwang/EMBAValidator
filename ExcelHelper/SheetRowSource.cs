using System;
using System.Collections.Generic;
using System.Linq;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// 代表Excel當中的資料列
    /// </summary>
    public class SheetRowSource : IRowStream,ICloneable
    {
        /// <summary>
        /// 建構式，傳入解析Excel物件
        /// </summary>
        /// <param name="Sheet"></param>
        public SheetRowSource(SheetHelper Sheet)
        {
            this.Sheet = Sheet;
            _fields = Sheet.Fields;
            CurrentRowIndex = -1;
        }

        /// <summary>
        /// 解析Excel的物件
        /// </summary>
        public SheetHelper Sheet { get; private set; }

        /// <summary>
        /// 取得或設定欄位名稱集合
        /// 設定：必須是Excel worksheet中既有的欄位，若不在當中則會忽略
        /// </summary>
        public List<string> Fields
        {
            get { return _fields; }
            set { _fields = Sheet.Fields.Intersect(value).ToList(); }
        }
        private List<string> _fields;

        /// <summary>
        /// 取得目前資料列索引
        /// </summary>
        public int CurrentRowIndex { get; private set; }

        /// <summary>
        /// 移至指定的資料列
        /// </summary>
        /// <param name="index">資料列索引</param>
        public void BindRow(int RowIndex)
        {
            CurrentRowIndex = RowIndex;
        }

        #region IRowStream 成員

        public string GetValue(string fieldName)
        {
            return Sheet.GetValue(CurrentRowIndex, Sheet.GetFieldIndex(fieldName)).Trim();
        }

        public bool Contains(string fieldName)
        {
            return Fields.Contains(fieldName);
        }

        public int Position
        {
            get { return CurrentRowIndex; }
        }

        #endregion

        #region IEnumerable<string> Members

        public IEnumerator<string> GetEnumerator()
        {
            return Fields.GetEnumerator();
        }

        #endregion

        #region IEnumerable Members

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return Fields.GetEnumerator();
        }

        #endregion

        #region ICloneable 成員

        public object Clone()
        {
            Dictionary<string,string> FieldValues = new Dictionary<string,string>();

            Fields.ForEach(x => FieldValues.Add(x, GetValue(x)));

            RowStream Row = new RowStream(FieldValues, Position);

            return Row;
        }

        #endregion
    }
}