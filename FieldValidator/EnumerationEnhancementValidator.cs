using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml;
using Aspose.Cells;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class EnumerationEnhancementValidator : IFieldValidator
    {
        private EnumerationValidator Validator = null;

        public EnumerationEnhancementValidator(XmlElement XmlNode)
        {
            //存放動態取得之列舉值清單
            List<string> dynamicValues = new List<string>();
            //靜態之列舉值清單 (直接寫在驗證規則裡)
            List<string> staticValues = new List<string>();

            foreach (XmlElement Elm in XmlNode.SelectNodes("Source"))
            {
                string Type = Elm.GetAttribute("Type");

                if (Type.Equals("Excel"))
                {
                    #region Excel
                    try
                    {
                        string Path = @Elm.GetAttribute("Path");
                        string SheetName = Elm.GetAttribute("SheetName");
                        string ColumnName = Elm.GetAttribute("ColumnName");
                        dynamicValues.AddRange(GetItemsFromExcel(Path, SheetName, ColumnName));
                    }
                    catch
                    {
                    }
                    #endregion
                }
                else if (Type.Equals("UserPreferences"))
                {
                    #region UserPreferences
                    try
                    {
                        string Name = Elm.GetAttribute("Name");
                        XmlDocument xmlDoc = new XmlDocument();

                        //是否有資料夾
                        if (!Directory.Exists(Path.Combine(Application.StartupPath, "ValidationReports")))
                        {
                            Directory.CreateDirectory(Path.Combine(Application.StartupPath, "ValidationReports"));
                        }

                        xmlDoc.Load(Path.Combine(Application.StartupPath + "\\ValidationReports", "ValidationPreferences.xml"));

                        XmlElement Mapping = (XmlElement)xmlDoc.DocumentElement.SelectSingleNode("Mapping[@Name='" + Name + "']");
                        if (Mapping != null)
                        {
                            dynamicValues.AddRange(GetItemsFromExcel(
                                Mapping.GetAttribute("Path"),
                                Mapping.GetAttribute("SheetName"),
                                Mapping.GetAttribute("ColumnName")
                                ));
                        }
                    }
                    catch
                    {
                    }
                    #endregion
                }
            }

            XmlDocument doc = new XmlDocument();
            XmlElement ContentXmlNode = doc.CreateElement("FieldValidator");

            //加入靜態的 Item Nodes
            foreach (XmlElement item in XmlNode.SelectNodes("Item"))
            {
                ContentXmlNode.AppendChild(ContentXmlNode.OwnerDocument.ImportNode(item, true));
                staticValues.Add(item.GetAttribute("Value"));
            }

            //加入動態的 Item Nodes
            foreach (string value in dynamicValues.Except(staticValues).ToList())
            {
                XmlElement ItemElm = ContentXmlNode.OwnerDocument.CreateElement("Item");
                ItemElm.SetAttribute("Value", value);
                ContentXmlNode.AppendChild(ItemElm);
            }

            Validator = new EnumerationValidator(ContentXmlNode);
        }

        private List<string> GetItemsFromExcel(string Path, string SheetName, string ColumnName)
        {
            List<string> items = new List<string>();

            if (!string.IsNullOrEmpty(Path) && !string.IsNullOrEmpty(SheetName) && !string.IsNullOrEmpty(ColumnName))
            {
                try
                {
                    Workbook book = new Workbook();

                    book.Open(Path);

                    for (int i = 0; i < book.Worksheets[SheetName].Cells.MaxDataColumn; i++)
                        if (book.Worksheets[SheetName].Cells[0, i].StringValue.Equals(ColumnName))
                        {
                            for (int j = 1; j <= book.Worksheets[SheetName].Cells.MaxDataRow; j++)
                                if (!string.IsNullOrEmpty(book.Worksheets[SheetName].Cells[j, i].StringValue))
                                {
                                    items.Add(book.Worksheets[SheetName].Cells[j, i].StringValue);

                                }
                                else
                                    break;
                            break;
                        }
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            return items;
        }

        #region IFieldValidator Members

        public bool Validate(string Value)
        {
            return Validator.Validate(Value);
        }

        public string Correct(string Value)
        {
            return Validator.Correct(Value);
        }

        public string ToString(string template)
        {
            return Validator.ToString(template);
        }

        #endregion
    }
}