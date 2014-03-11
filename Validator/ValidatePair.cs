
namespace EMBA.Validator
{
    /// <summary>
    /// 尚未驗證的驗證規則及驗證資料組合
    /// </summary>
    public class ValidatePair
    {
        /// <summary>
        /// 驗證規則路徑
        /// </summary>
        public string ValidatorFile { get; set; }

        /// <summary>
        /// 驗證File路徑
        /// </summary>
        public string DataFile { get; set; }

        /// <summary>
        /// 驗證Sheet名稱
        /// </summary>
        public string DataSheet { get; set; }

        /// <summary>
        /// 空白建構式
        /// </summary>
        public ValidatePair()
        {
 
        }

        /// <summary>
        /// 建構式，傳入驗證檔案、來源Excel檔案路徑及資料表
        /// </summary>
        /// <param name="ValidatorFile"></param>
        /// <param name="DataFile"></param>
        /// <param name="DataSheet"></param>
        public ValidatePair(string ValidatorFile, string DataFile, string DataSheet)
            : this()
        {
            this.ValidatorFile = ValidatorFile;
            this.DataFile = DataFile;
            this.DataSheet = DataSheet;
        }
    }
}