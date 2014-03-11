using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    class AppellationValidator : IRowVaildator
    {

        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            if (Value.Contains("監護人關係") && Value.Contains("監護人稱謂"))
                if (string.IsNullOrEmpty(Value.GetValue("監護人稱謂")) && !string.IsNullOrEmpty(Value.GetValue("監護人關係")))
                    return false;
            
            return true;

            //return !string.IsNullOrEmpty(Value.GetValue("監護人稱謂"));
        }

        public string Correct(IRowStream Value)
        {
            if (Value.Contains("監護人關係") && Value.Contains("監護人稱謂"))
            {
                if (string.IsNullOrEmpty(Value.GetValue("監護人稱謂")) && !string.IsNullOrEmpty(Value.GetValue("監護人關係")))
                {
                    return "<A><監護人稱謂>" + Value.GetValue("監護人關係") + "</監護人稱謂></A>";
                }
            }

            return "<A><監護人稱謂/></A>";
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion
    }
}
