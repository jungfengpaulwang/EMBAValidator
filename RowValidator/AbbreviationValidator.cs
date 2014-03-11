using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class AbbreviationValidator : IRowVaildator
    {
        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            return !string.IsNullOrEmpty(Value.GetValue("縮寫"));
        }

        public string Correct(IRowStream Value)
        {
            if (Value.Contains("缺曠名稱"))
            {
                string name = Value.GetValue("缺曠名稱");
                if (!string.IsNullOrEmpty(name))
                {
                    return "<A><縮寫>" + name.Substring(0, 1) + "</縮寫></A>";
                }
                else
                {
                    return string.Empty;
                }
            }

            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}
