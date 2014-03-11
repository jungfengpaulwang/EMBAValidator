using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Validator
{
    public class FalseRowValidator : EMBA.DocumentValidator.IRowVaildator
    {
        #region IRowVaildator 成員

        public bool Validate(EMBA.DocumentValidator.IRowStream Value)
        {
            return false;
        }

        public string Correct(EMBA.DocumentValidator.IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}
