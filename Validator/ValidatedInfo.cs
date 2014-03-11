using System.Collections.Generic;
using Aspose.Cells;
using System.IO;

namespace EMBA.Validator
{
    /// <summary>
    /// 驗證完後的訊息
    /// </summary>
    public class ValidatedInfo
    {
        private Workbook mResult;
        private SheetHelper mResultHelper;

        /// <summary>
        /// 存檔位置
        /// </summary>
        public string OutputFile { get ; set;}

        /// <summary>
        /// 產生驗證報告方式
        /// </summary>
        public OutputOptions OutputOptions { get; set; }

        /// <summary>
        /// 驗證後的訊息
        /// </summary>
        public List<ValidatedPair> ValidatedPairs { get; set; }

        /// <summary>
        /// 取得預設的DataFile路徑，為ValidatedPairs中的第一個DataFile
        /// </summary>
        public string DefaultDataFile
        {
            get
            {
                if (ValidatedPairs != null && ValidatedPairs.Count > 0)
                    return ValidatedPairs[0].DataFile;

                return string.Empty;
            }
        }

        /// <summary>
        /// 根據DefaultDataFile取得檔案名稱
        /// </summary>
        public string DefaultDataFileName
        {
            get
            {
                if (!string.IsNullOrEmpty(DefaultDataFile))
                    return (new FileInfo(DefaultDataFile)).Name;

                return string.Empty;
            }
        }

        /// <summary>
        /// 驗證完後的實際資料
        /// </summary>
        public Workbook Result 
        {
            get
            {
                return mResult;
            }
            set
            {
                mResult = value;
            }
        }

        /// <summary>
        /// 驗證完後的資料可用SheetHelper讀取
        /// </summary>
        public SheetHelper ResultHelper  
        { 
            get 
            {
                return mResultHelper;
            }
            set
            {
                mResultHelper = value;
            }           
        }

        /// <summary>
        /// 預設建構式
        /// </summary>
        public ValidatedInfo()
        {
            ValidatedPairs = new List<ValidatedPair>();
            OutputFile = string.Empty;
        }
    }
}