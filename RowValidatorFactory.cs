using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class RowValidatorFactory : IRowValidatorFactory
    {
        #region IRowValidatorFactory 成員

        public IRowVaildator CreateRowValidator(string typeName, System.Xml.XmlElement validatorDescription)
        {
            switch (typeName.ToUpper())
            {
                case "ENROLLYEARMONTH":
                    return new EnrollYearMonthRowValidator();
                case "UPDATEDATE":
                    return new UpdateDateRowValidator();
                case "SERIANUMBER":
                    return new SerialNumberValidator();
                case "ABBREVIATION":
                    return new AbbreviationValidator();
                case "HOTKEY":
                    return new HotKeyValidator();
                case "APPELLATION":
                    return new AppellationValidator();
                case "ATTENDENCESCHOOLYEARSEMESTER":
                    return new AttendenceSchoolYearSemesterValidator();
                case "DISCIPLINESCHOOLYEARSEMESTER":
                    return new DisciplineSchoolYearSemesterValidator();
                case "UPDATEPERMRECREASON":
                    return new UpdatePermrecReasonValidator();
                case "":
                    return new FalseRowValidator();
                default:
                    return null;
            }
        }

        #endregion
    }
}
