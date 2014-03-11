using System;
using System.Collections.Generic;
using System.Linq;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// 已驗證完的驗證規則及驗證資料組合
    public class ValidatedPair : ValidatingPair
    {
        /// <summary>
        /// 建構式
        /// </summary>
        public ValidatedPair()
        {
            Rows = new List<IRowStream>();
        }

        /// <summary>
        /// 實際抽象化資料列
        /// </summary>
        public List<IRowStream> Rows { get; set; }

        /// <summary>
        /// 經過欄位驗證後的資訊
        /// </summary>
        public IList<FieldValidatedDescription> FieldDescriptions { get; set; }

        /// <summary>
        /// 資料重覆驗證結果
        /// </summary>
        public IList<DuplicateData> Duplicates { get; set; }

        /// <summary>
        /// 驗證過程中是否有錯誤
        /// </summary>
        public IList<Exception> Exceptions { get; set; }

        /// <summary>
        /// 判斷是否可匯入，其條件為必填欄位都有，並且錯誤數目為0
        /// </summary>
        public bool Importable
        {
            get
            {
                if (Exceptions != null && Exceptions.Count > 0)
                    return false;

                //若錯誤數量為0，則回傳false
                if (ErrorCount > 0)
                    return false;

                //取得必填欄位但是在資料來源中沒有的欄位
                IEnumerable<FieldValidatedDescription> NotInSourceRequiredFields = FieldDescriptions.Where(x => x.InDefinition && x.Required && !x.InSource);

                //若是欄位的數量大於0，則回傳false
                if (NotInSourceRequiredFields != null)
                    if (NotInSourceRequiredFields.Count() > 0)
                        return false;

                //判斷主鍵資料是否有重覆，而且是錯誤型態
                foreach (DuplicateData Duplicate in Duplicates)
                    if (Duplicate.ErrorType == EMBA.DocumentValidator.ErrorType.Error)
                        if (Duplicate.Count > 0)
                            return false;

                //若上述兩個條件皆不成立，則回傳true
                return true;
            }
        }

        /// <summary>
        /// 更新訊息
        /// </summary>
        public void UpdateMessage()
        {
            Message = "錯誤數目：" + ErrorCount + "、警告數目：" + WarningCount + "、自動修正數目：" + AutoCorrectCount + Environment.NewLine;
            Message += FieldDescriptions.ToDisplay();
        }
    }
}