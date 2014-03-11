
namespace EMBA.Validator
{
    /// <summary>
    /// 驗證中的驗證規則及驗證資料組合
    /// </summary>
    public class ValidatingPair : ValidatePair
    {
        /// <summary>
        /// 指定訊息
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 錯誤計數
        /// </summary>
        public int ErrorCount { get; set; }

        /// <summary>
        /// 警告計數
        /// </summary>
        public int WarningCount { get; set; }

        /// <summary>
        /// 自動修正計數
        /// </summary>
        public int AutoCorrectCount { get; set; }

        ///// <summary>
        ///// 檔案名稱
        ///// </summary>
        //public string DataFile { get; set; }

        ///// <summary>
        ///// 工作表名稱
        ///// </summary>
        //public string DataSheet { get; set; }
    }
}