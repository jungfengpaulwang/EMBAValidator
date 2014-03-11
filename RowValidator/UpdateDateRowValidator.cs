using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// 只處理異動日期
    /// 自動修正：(學年度+1911)/8/31
    /// </summary>
    public class UpdateDateRowValidator : IRowVaildator
    {
        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            return !string.IsNullOrEmpty(Value.GetValue("異動日期"));
        }

        public string Correct(IRowStream Value)
        {
            int i;
            if (int.TryParse(Value.GetValue("學年度"), out i))
                return "<A><異動日期>" + (i + 1911) + "/8/31" + "</異動日期></A>";
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
