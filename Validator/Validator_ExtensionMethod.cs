using System;
using System.Xml;
using EMBA.DocumentValidator;
using FISCA.Presentation.Controls;

namespace EMBA.Validator
{
    public static class Validator_ExtensionMethod
    {
        /// <summary>
        /// 從檔案載入驗證規則
        /// </summa>
        /// <param name="DocumentValidate"></param>
        /// <param name="Filename"></param>
        public static void Load(this DocumentValidate docValidate, string Filename)
        {
            XmlDocument xmldoc = new XmlDocument();

            try
            {
                //用XmlDocument物件載入驗證規則
                xmldoc.Load(Filename);
                //Validator載入驗證規則
                docValidate.LoadRule(xmldoc.DocumentElement);
            }
            catch (Exception ex)
            {
                //MsgBox.Show("讀取驗證規則路徑發生錯誤!\n\n" + ex.Message);
                //SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
            }
        }
    }
}