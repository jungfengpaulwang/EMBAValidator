using System;
using System.Collections.Generic;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// 只處理入學年月
    /// 自動修正：(學年度+1911) + "09"
    /// </summary>
    public class EnrollYearMonthRowValidator : IRowVaildator
    {
        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            return !string.IsNullOrEmpty(Value.GetValue("入學年月"));
        }

        public string Correct(IRowStream Value)
        {
            int i;
            if (int.TryParse(Value.GetValue("學年度"), out i))
                return "<A><入學年月>" + (i + 1911) + "09" + "</入學年月></A>";
            else
                return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}
