using System;
using System.Collections.Generic;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// IRowStream的簡單實作，傳入Dictionary及Position即可運作
    /// </summary>
    public class RowStream : IRowStream
    {
        private Dictionary<string, string> mFieldValues;
        private int mPosition;

        /// <summary>
        /// 建構式，傳入Dictionary及Position
        /// </summary>
        /// <param name="FieldValues"></param>
        /// <param name="Position"></param>
        public RowStream(Dictionary<string, string> FieldValues, int Position)
        {
            if (FieldValues == null)
                throw new ArgumentNullException("傳入的參數FieldValues不得為空白!");

            this.mFieldValues = FieldValues;
            this.mPosition = Position;
        }

        #region IRowStream 成員

        /// <summary>
        /// 根據欄位名稱取得值
        /// </summary>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public string GetValue(string fieldName)
        {
            return mFieldValues.ContainsKey(fieldName) ? mFieldValues[fieldName] : string.Empty;
        }

        /// <summary>
        /// 是否包含某欄位
        /// </summary>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public bool Contains(string fieldName)
        {
            return mFieldValues.ContainsKey(fieldName);
        }

        /// <summary>
        /// 在來源資料的所在位置
        /// </summary>
        public int Position
        {
            get { return mPosition ; }
        }

        #endregion

        #region IEnumerable<string> 成員

        public IEnumerator<string> GetEnumerator()
        {
            return mFieldValues.Keys.GetEnumerator();
        }

        #endregion

        #region IEnumerable 成員

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return mFieldValues.Keys.GetEnumerator();
        }

        #endregion
    }
}