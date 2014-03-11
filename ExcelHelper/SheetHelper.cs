using System;
using System.Collections.Generic;
using System.Drawing;
using Aspose.Cells;

namespace EMBA.Validator
{
    /// <summary>
    /// 協助讀取資料表的類別
    /// </summary>
    public class SheetHelper
    {
        private const int StartRow = 0, StartColumn = 0;

        private Workbook _book;
        private Worksheet _sheet;
        private Dictionary<string, SheetField> _fields;

        /// <summary>
        /// 建構式，傳入Excel檔案以及指定的Sheet名稱，會將此Excel設為預設值，並切換到指定的Sheet。
        /// </summary>
        /// <param name="sourceBook">Workbook物件</param>
        /// <param name="sheetName">Worksheet名稱</param>
        public SheetHelper(Workbook sourceBook, string sheetName)
        {
            Initial(sourceBook, sheetName);
        }

        /// <summary>
        /// 建構式，傳入Excel檔案路徑以及指定的Sheet名稱，會開啟此Excel檔案，並切換到指定的Sheet。
        /// </summary>
        /// <param name="sourceFile">Excel檔案路徑</param>
        /// <param name="sheetName">指定的Sheet名稱</param>
        public SheetHelper(string sourceFile,string sheetName)
        {
            Initial(sourceFile, sheetName);
        }

        public SheetHelper(string sourceFile)
        {
            Initial(sourceFile, string.Empty);
        }

        private void Initial(Workbook sourceBook, string sheetName)
        {
            //若Workbook物件為null則丟出錯誤
            if (sourceBook == null)
                throw new NullReferenceException("SheetHelper錯誤，不能傳入null的Workbook物件。");

            _book = sourceBook;
            _sheet = GetWorksheet(_book, sheetName);
            _fields = GetFieldList(_sheet);
        }

        private void Initial(string sourceFile,string sheetName)
        {
            _book = GetWorkbook(sourceFile);
            _sheet = GetWorksheet(_book, sheetName);
            _fields = GetFieldList(_sheet);
        }

        /// <summary>
        /// 切換資料表
        /// </summary>
        /// <param name="sheetName">資料表名稱</param>
        public void SwitchSeet(string sheetName)
        {
            _sheet = GetWorksheet(_book, sheetName);
            _fields = GetFieldList(_sheet);
        }

        /// <summary>
        /// 取得作用中的活頁簿
        /// </summary>
        public Workbook Book
        {
            get { return _book; }
        }

        /// <summary>
        /// 取得作用中的資料表
        /// </summary>
        public Worksheet Sheet
        {
            get { return _sheet; }
        }

        /// <summary>
        /// 根據欄位名稱取得在資料表中的順序
        /// </summary>
        /// <param name="fieldName">欄位名稱</param>
        /// <returns>順序</returns>
        public int GetFieldIndex(string fieldName)
        {
            if (_fields.ContainsKey(fieldName))
                return _fields[fieldName].AbsoluteIndex;
            else
                return -1;
        }

        /// <summary>
        /// 資料表中的欄位名稱列表
        /// </summary>
        public List<string> Fields
        {
            get
            {
                return new List<string>(_fields.Keys);
            }
        }

        /// <summary>
        /// 根據顏色來取得欄位名稱
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public List<string> GetFieldsByColor(Color color)
        {
            List<string> fields = new List<string>();
            foreach (string each in Fields)
            {
                Style fStyle = _sheet.Cells[StartRow, GetFieldIndex(each)].Style;

                Color c1 = MatchColor(fStyle.ForegroundColor);
                Color c2 = MatchColor(color);

                if (c1 == c2) fields.Add(each);
            }

            return fields;
        }

        /// <summary>
        /// 設定多個欄位的格式
        /// </summary>
        /// <param name="fields"></param>
        /// <param name="style"></param>
        public void SetFieldsStyle(List<string> fields, Style style)
        {
            foreach (string each in fields)
                _sheet.Cells[StartRow, GetFieldIndex(each)].Style = style;
        }

        /// <summary>
        /// 設定儲存格值。*
        /// </summary>
        /// <param name="row">zero based.</param>
        /// <param name="column">zero based.</param>
        /// <returns></returns>
        public string GetValue(int row, int column)
        {
            try
            {
                return _sheet.Cells[row, column].StringValue;
            }
            catch (ArgumentException ex)
            {
                return string.Empty;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 設定儲存格值。
        /// </summary>
        /// <param name="row">zero based.</param>
        /// <param name="column">zero based.</param>
        /// <returns></returns>
        public void SetValue(int row, int column, string value)
        {
            _sheet.Cells[row, column].PutValue(value);
        }

        public void SetStyle(int row, int column, Style style)
        {
            _sheet.Cells[row, column].Style = style;
        }

        /// <summary>
        /// 將所有含有資料的儲存格設成統一的樣式。
        /// </summary>
        /// <param name="style"></param>
        public void SetAllStyle(Style style)
        {
            int maxColumn = _sheet.Cells.MaxDataColumn;
            int maxRow = MaxDataRowIndex;

            Range rng = _sheet.Cells.CreateRange(0, 0, maxRow + 1, maxColumn + 1);
            rng.Style = style;
        }

        /// <summary>
        /// 設定某個儲存格的註解
        /// </summary>
        /// <param name="row">資料列</param>
        /// <param name="column">資料行</param>
        /// <param name="msg">註解內容</param>
        public void SetComment(int row, int column, string msg)
        {
            Comment cmm = _sheet.Comments[_sheet.Comments.Add(row, (byte)column)];
            cmm.Note = msg;
        }

        /// <summary>
        /// 清除資料表中的所有註解
        /// </summary>
        public void ClearComments()
        {
            _sheet.Comments.Clear();
        }

        /// <summary>
        /// 建立一個新的樣式。
        /// </summary>
        /// <returns></returns>
        public Style NewStyle()
        {
            return Book.Styles[Book.Styles.Add()];
        }

        public Color MatchColor(Color color)
        {
            return Book.GetMatchingColor(color);
        }

        //public void Save(string fileName)
        //{
        //    _book.Save(fileName);
        //    Initial(fileName);
        //}

        /// <summary>
        /// 第一個有資料的列索引
        /// </summary>
        public int FirstDataRowIndex { get { return StartRow + 1; } }

        /// <summary>
        /// 最大資料列索引
        /// </summary>
        public int MaxDataRowIndex
        {
            get { return _sheet.Cells.MaxDataRow; }
        }

        /// <summary>
        /// 資料列個數，去除標題列
        /// </summary>
        public int DataRowCount
        {
            get { return _sheet.Cells.MaxDataRow - StartRow; }
        }

        /// <summary>
        /// 根據活頁簿物件及資料表名稱取得對應的資料表物件參照
        /// </summary>
        /// <param name="book">活頁簿物件</param>
        /// <param name="sheetName">資料表名稱</param>
        /// <returns>對應資料表名稱的資料表物件</returns>
        private static Worksheet GetWorksheet(Workbook book,string sheetName)
        {
            Worksheet sheet = null;

            try
            {
                if (string.IsNullOrEmpty(sheetName))
                    sheet = book.Worksheets[0];
                else
                    sheet = book.Worksheets[sheetName];
            }
            catch (Exception e)
            {
                throw new ArgumentException("讀取工作表資訊失敗。", e);
            }
            return sheet;
        }

        /// <summary>
        /// 根據檔案路徑取得活頁簿物件
        /// </summary>
        /// <param name="sourceFile">檔案路徑</param>
        /// <returns>活頁簿物件</returns>
        private static Workbook GetWorkbook(string sourceFile)
        {
            Workbook book = new Workbook();
            book.Open(sourceFile);

            return book;
        }

        /// <summary>
        /// 根據資料表讀取欄位定義
        /// </summary>
        /// <param name="sheet">資料表</param>
        /// <returns></returns>
        private static Dictionary<string, SheetField> GetFieldList(Worksheet sheet)
        {
            Cell startCell = sheet.Cells[StartRow, StartColumn];
            int maxColumn = sheet.Cells.MaxDataColumn; //最大的資料欄 Index (zero based)。
            Dictionary<string, SheetField> columns = new Dictionary<string, SheetField>();

            for (int i = StartColumn; i <= maxColumn; i++)
            {
                Cell colCell = sheet.Cells[StartRow, i];
                if (colCell.StringValue != null && colCell.StringValue.Trim() != string.Empty)
                {
                    SheetField objField = new SheetField(colCell);

                    if (columns.ContainsKey(objField.FieldName))
                        throw new ArgumentException("重覆的欄位名稱：" + objField.FieldName);

                    columns.Add(objField.FieldName, objField);
                }
            }

            return columns;
        }

        #region InnerClass SheetField
        private class SheetField
        {
            public SheetField(Cell fieldCell)
            {
                FieldName = fieldCell.StringValue;
                AbsoluteIndex = fieldCell.Column;
            }

            /// <summary>
            /// 欄位名稱
            /// </summary>
            public string FieldName { get; set;}

            /// <summary>
            /// 欄位索引，以0為開始
            /// </summary>
            public int AbsoluteIndex { get; set;}
        }
        #endregion
    }
}