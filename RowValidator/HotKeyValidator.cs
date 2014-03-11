using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class HotKeyValidator : IRowVaildator
    {
        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            return !string.IsNullOrEmpty(Value.GetValue("熱鍵"));
        }

        public string Correct(IRowStream Value)
        {
            int Position = Value.Position - 1;

            string Result = Position > 9 ? "" + Convert.ToChar(Position + 55) : "" + Position;

            return "<A><熱鍵>" + Result + "</熱鍵></A>";
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}
