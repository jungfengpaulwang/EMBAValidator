using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// 處理學籍異動中的更正學籍自動修正
    /// </summary>
    public class UpdatePermrecReasonValidator : IRowVaildator
    {

        #region IRowVaildator Members

        private bool IsCorrectable(IRowStream Value)
        {
            if (Value.Contains("新生日") && !string.IsNullOrEmpty(Value.GetValue("新生日")))
                return true;

            if (Value.Contains("新身分證號") && !string.IsNullOrEmpty(Value.GetValue("新身分證號")))
                return true;

            if (Value.Contains("新姓名") && !string.IsNullOrEmpty(Value.GetValue("新姓名")))
                return true;

            if (Value.Contains("新性別") && !string.IsNullOrEmpty(Value.GetValue("新性別")))
                return true;

            return false;
        }

        public bool Validate(IRowStream Value)
        {
            //1.判斷欄位資料來源是否包含『異動類別』及『原因及事項』欄位
            //2.判斷『異動類別』為更正學籍，以及『原因及事項』為空白。
            //3.判斷以下欄位其中一個『新生日』、『新身分證號』、『新姓名』、『新性別』有包含欄位，而且不為空白。
            //3.符合1、2點就傳回false，否則傳回true
            if (Value.Contains("異動類別") && Value.Contains("原因及事項"))
                if (Value.GetValue("異動類別").Equals("更正學籍") && string.IsNullOrEmpty(Value.GetValue("原因及事項")))
                    if (IsCorrectable(Value))
                        return false;
         
            return true;
        }

        public string Correct(IRowStream Value)
        {
            if (Value.Contains("異動類別") && Value.Contains("原因及事項"))
                if (Value.GetValue("異動類別").Equals("更正學籍") && string.IsNullOrEmpty(Value.GetValue("原因及事項")))
                {
                    string Result = string.Empty;

                    if (Value.Contains("新生日") && !string.IsNullOrEmpty(Value.GetValue("新生日")))
                        Result += "新生日、";

                    if (Value.Contains("新身分證號") && !string.IsNullOrEmpty(Value.GetValue("新身分證號")))
                        Result += "新身分證號、";

                    if (Value.Contains("新姓名") && !string.IsNullOrEmpty(Value.GetValue("新姓名")))
                        Result += "新姓名、";

                    if (Value.Contains("新性別") && !string.IsNullOrEmpty(Value.GetValue("新性別")))
                        Result += "新性別、";

                    if (!string.IsNullOrEmpty(Result))
                        Result = "更正" + Result.Substring(0, Result.Length - 1);

                    Result = "<A><原因及事項>" + Result + "</原因及事項></A>";

                    return Result;
                }

            return "<A><原因及事項/></A>";
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}