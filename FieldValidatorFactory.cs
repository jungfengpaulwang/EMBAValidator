using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class FieldValidatorFactory : IFieldValidatorFactory
    {
        #region IFieldValidatorFactory 成員

        public IFieldValidator CreateFieldValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            //foreach(Assembly Assembly in AppDomain.CurrentDomain.GetAssemblies())
            //    foreach(Type Type in Assembly.GetTypes())
            //        foreach(Type Interface in Type.GetInterfaces())
            //            if (Interface is typeof(DocValidate.IFieldValidator))
            //                {}

            switch (typeName.ToUpper())
            {
                case "ENUMERATIONENHANCEMENT":
                    return new EnumerationEnhancementValidator(validatorDescription);
                default:
                    return null;
            }
        }

        #endregion
    }
}
