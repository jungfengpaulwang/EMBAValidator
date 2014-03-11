using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    /// <summary>
    /// 進度回報宣告
    /// </summary>
    /// <param name="Pair"></param>
    /// <param name="ErrorCount"></param>
    /// <param name="WarningCount"></param>
    /// <param name="AutoCorrectCount"></param>
    /// <param name="Message"></param>
    /// <param name="CurrentProgress"></param>
    public delegate void ReportProgress(ValidatePair Pair, int ErrorCount, int WarningCount, int AutoCorrectCount, string Message, int CurrentProgress);

    /// <summary>
    /// 簡化後的資料驗證器，可直接傳入多組驗證規則及Excel檔案進行驗證
    /// </summary>
    public class Validator
    {
        /// <summary>
        /// 驗證進度
        /// </summary>
        public Action<ValidatingPair, int> Progress;

        /// <summary>
        /// 自行驗證
        /// </summary>
        public Action<List<IRowStream>, RowMessages> CustomValidate;

        /// <summary>
        /// 驗證完成訊息
        /// </summary>
        public Action<ValidatedInfo> Complete;

        /// <summary>
        /// 驗證中的訊息記錄
        /// </summary>
        private RowMessages RowMessages;

        /// <summary>
        /// 錯誤數量
        /// </summary>
        private int ErrorCount;

        /// <summary>
        /// 警告數量
        /// </summary>
        private int WarningCount;
        
        /// <summary>
        /// 自動修正數量
        /// </summary>
        private int AutoCorrectCount;

        /// <summary>
        /// 建構式
        /// </summary>
        public Validator()
        {
            RowMessages = new RowMessages();
        }

        /// <summary>
        /// 資料驗證
        /// </summary>
        /// <param name="ValidatorFile">驗證規則路徑</param>
        /// <param name="DataFile">驗證檔案路徑</param>
        /// <param name="DataSheet">驗證的工作表</param>
        /// <param name="OutputFile">報告輸出路徑</param>
        public void Validate(string ValidatorFile, string DataFile, string DataSheet, string OutputFile)
        {
            ValidatePair validatorPair = new ValidatePair(ValidatorFile, DataFile, DataSheet);
            Validate(new ValidatePair[] { validatorPair }, OutputFile, OutputOptions.Full);
        }

        /// <summary>
        /// 資料驗證
        /// </summary>
        /// <param name="ValidatePair">資料驗證相關資訊(檔案,工作表,規則)</param>
        /// <param name="OutputFile">報告輸出路徑</param>
        public void Validate(ValidatePair ValidatePair, string OutputFile)
        {
            Validate(new ValidatePair[] { ValidatePair }, OutputFile, OutputOptions.Full);
        }

        /// <summary>
        /// 資料驗證
        /// </summary>
        /// <param name="ValidatorPairs">資料驗證相關資訊清單(檔案,工作表,規則)</param>
        /// <param name="OutputFile">報告輸出路徑</param>
        public void Validate(IEnumerable<ValidatePair> ValidatorPairs, string OutputFile)
        {
            Validate(ValidatorPairs, OutputFile, OutputOptions.Full);
        }

        /// <summary>
        /// 資料驗證
        /// </summary>
        /// <param name="ValidatorPairs">資料驗證相關資訊清單(檔案,工作表,規則)</param>
        /// <param name="OutputFile">報告輸出路徑</param>
        /// <param name="OutputOptions">輸出類型:全部(Full),正確(Correct),錯誤(Error)</param>
        protected void Validate(IEnumerable<ValidatePair> ValidatorPairs, string OutputFile, OutputOptions OutputOptions)
        {
            #region Step1：初始化動作
            //假設ValidatorPairs為0的話就不處理
            if (ValidatorPairs.Count() == 0)
                return;

            if (!OutputBuilder.SameSource(ValidatorPairs,OutputOptions))
                throw new Exception("目前驗證只支援ValidatePair中的DataFile需為一樣，並且OutputOptions為Full。");

            //建立OutputBuilder物件，用來協助產生最後的Excel檔案
            OutputBuilder outputBuilder = new OutputBuilder(ValidatorPairs, OutputFile, OutputOptions,this.ReportProgress);

            //建立ValidatedInfo物件，用來儲存資料驗證後的結果
            ValidatedInfo validatedInfo = new ValidatedInfo();

            //初始化DocumentValidate物件及事件
            DocumentValidate docValidate = new DocumentValidate();
            docValidate.AutoCorrect += new EventHandler<AutoCorrectEventArgs>(docValidate_AutoCorrect);
            docValidate.ErrorCaptured += new EventHandler<ErrorCapturedEventArgs>(docValidate_ErrorCaptured);
            #endregion

            #region Step2：進行資料驗證

            foreach (ValidatePair each in ValidatorPairs)
            {
                ValidatedPair validatedPair = new ValidatedPair();

                //初始化驗證訊息，在驗證每個資料表時會重新清空
                RowMessages = new RowMessages();

                //用來搜集資料驗證過程中的錯誤訊息, 等到資料驗證完後再一次回報錯誤訊息
                Dictionary<string, Exception> exceptions = new Dictionary<string, Exception>();

                List<FieldValidatedDescription> FieldDescriptions = new List<FieldValidatedDescription>();

                IList<DuplicateData> DuplicateDataList = new List<DuplicateData>();

                try
                {
                    #region Step2.1：讀取驗證規則
                    docValidate.Load(each.ValidatorFile);

                    #endregion

                    #region Step2.2：驗證資料
                    //將驗證訊息都清空
                    RowMessages.Clear();

                    //將錯誤、警告及自動修正數目設為0
                    ErrorCount = WarningCount = AutoCorrectCount = 0;

                    //根據目前的驗證資料組合切換
                    outputBuilder.Switch(each);

                    //開始進度回報
                    ReportProgress(each, ErrorCount, WarningCount, AutoCorrectCount, "目前驗證: " + each.DataSheet, 0);


                    double ValidateProgress = 0;
                    double ValidateTotal = outputBuilder.Sheet.DataRowCount;

                    //開始主鍵驗證
                    docValidate.BeginDetecteDuplicate();

                    #region Step2.2.1：資料驗證
                    SheetRowSource RowSource = new SheetRowSource(outputBuilder.Sheet);
                    for (int i = outputBuilder.Sheet.FirstDataRowIndex; i <= outputBuilder.Sheet.DataRowCount; i++)
                    {
                        try
                        {
                            RowSource.BindRow(i);

                            bool valid = docValidate.Validate(RowSource);

                            RowStream RowStream = RowSource.Clone() as RowStream; //將SheetRowSource目前所指向的內容複製

                            validatedPair.Rows.Add(RowStream);                   //將RowStream加入到集合中

                            ReportProgress(each, ErrorCount, WarningCount, AutoCorrectCount, "目前驗證: " + each.DataSheet, (int)(++ValidateProgress * 100 / ValidateTotal));
                        }
                        catch (Exception e)
                        {
                            ReportProgress(each, ErrorCount, WarningCount, AutoCorrectCount, "驗證錯誤: " + each.DataSheet, (int)(++ValidateProgress * 100 / ValidateTotal));

                            if (!exceptions.ContainsKey(e.Message))
                                exceptions.Add(e.Message, e);
                        }
                    }

                    //exceptions.Values.ToList().ForEach(ex => SmartSchool.ErrorReporting.ReportingService.ReportException(ex));
                    //需要將驗證有錯誤呈現給使用者
                    #endregion

                    #region Step2.2.2：主鍵驗證
                    DuplicateDataList = docValidate.EndDetecteDuplicate();

                    foreach (DuplicateData dup in DuplicateDataList)
                    {
                        ErrorType errorType = (dup.ErrorType == EMBA.DocumentValidator.ErrorType.Error) ? ErrorType.Error : ErrorType.Warning;

                        foreach (DuplicateRecord dupRecord in dup)
                        {
                            List<string> list = new List<string>();
                            for (int i = 0; i < dupRecord.Values.Count; i++)
                                list.Add(string.Format("「{0}：{1}」", dup.Fields[i], string.IsNullOrWhiteSpace(dupRecord.Values[i]) ? "缺" : dupRecord.Values[i]));

                            string fields = (list.Count == 1 ? list[0] : string.Join("、", list.ToArray()));
                            string msg = string.Format("{0}為鍵值，必須有資料且不得重覆。", fields);
                            foreach (int position in dupRecord.Positions)
                                RowMessages[position].MessageItems.Add(new MessageItem(errorType, ValidatorType.Row, msg));
                        }
                    }
                    #endregion

                    #region Step2.2.3：欄位驗證
                    RowSource.BindRow(0);

                    FieldDescriptions = docValidate.ValidateField(RowSource);

                    string HeaderMessage = FieldDescriptions.ToDisplay();
                    #endregion

                    #region Step2.2.4:客製驗證
                    if (CustomValidate != null)
                        CustomValidate.Invoke(validatedPair.Rows, RowMessages);
                    #endregion

                    #endregion

                    #region Step2.3：將驗證訊息寫到來源Excel
                    outputBuilder.InitialMessageHeader();

                    //判斷當沒有訊息時就跳過不輸出
                    //if (RowMessages.Count <= 0 && string.IsNullOrEmpty(HeaderMessage)) continue;

                    //寫入欄位驗證訊息
                    outputBuilder.SetHeaderMessage(HeaderMessage);

                    //寫入資料及主鍵驗證訊息
                    outputBuilder.SetMessages(RowMessages);

                    //調整驗證訊息欄寬
                    outputBuilder.AutoFitMessage();
                    #endregion
                }
                catch (Exception e)
                {
                    if (!exceptions.ContainsKey(e.Message))
                        exceptions.Add(e.Message, e); 
                }

                #region Step2.4：儲存驗證結果
                validatedPair.AutoCorrectCount = RowMessages.AutoCorrectCount;
                validatedPair.ErrorCount = RowMessages.ErrorCount;
                validatedPair.WarningCount = RowMessages.WarningCount;
                validatedPair.DataFile = each.DataFile;
                validatedPair.DataSheet = each.DataSheet;
                validatedPair.ValidatorFile = each.ValidatorFile;
                validatedPair.FieldDescriptions = FieldDescriptions;
                validatedPair.Duplicates = DuplicateDataList;
                validatedPair.Exceptions = exceptions.Values.ToList();
                validatedPair.UpdateMessage();

                validatedInfo.ValidatedPairs.Add(validatedPair);
                #endregion
            }
            #endregion

            #region Step3：驗證報告存檔
            validatedInfo.OutputFile = OutputFile;
            validatedInfo.OutputOptions = OutputOptions;
            validatedInfo.Result = outputBuilder.Sheet.Book;
            validatedInfo.ResultHelper = outputBuilder.Sheet;

            outputBuilder.Save();
            #endregion

            #region 儲存驗證結果
            if (Complete != null)
                Complete(validatedInfo);
            #endregion
        }

        /// <summary>
        /// 回報驗證中的訊息
        /// </summary>
        /// <param name="Pair">驗證資料及規則組合</param>
        /// <param name="ErrorCount">錯誤數目</param>
        /// <param name="WarningCount">警告數目</param>
        /// <param name="AutoCorrectCount">自動修正數目</param>
        /// <param name="Message">進度回報訊息</param>
        /// <param name="CurrentProgress">目前進度</param>
        internal void ReportProgress(ValidatePair Pair,int ErrorCount,int WarningCount,int AutoCorrectCount,string Message,int CurrentProgress)
        {
            if (Progress != null)
            {
                ValidatingPair obj = new ValidatingPair();

                obj.ErrorCount = ErrorCount;
                obj.WarningCount = WarningCount;
                obj.AutoCorrectCount = AutoCorrectCount;

                obj.Message = Message;
                obj.DataFile = Pair != null ? Path.GetFileName(Pair.DataFile) : string.Empty;
                obj.DataSheet = Pair != null ? Pair.DataSheet : string.Empty;
                Progress(obj, CurrentProgress);
            }
        }

        /// <summary>
        /// 自動修正事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void docValidate_AutoCorrect(object sender, AutoCorrectEventArgs e)
        {
            #region 將修正後的值寫入該欄位
            SheetRowSource row = e.Row as SheetRowSource;
            SheetHelper sheetHelper = row.Sheet;

            int FieldIndex = sheetHelper.GetFieldIndex(e.FieldName);
            if (FieldIndex != -1) //當自動修正,Excel-Sheet內卻沒有欄位
            {
                sheetHelper.Sheet.Cells[row.Position, sheetHelper.GetFieldIndex(e.FieldName)].PutValue(e.NewValue);

                string Description = string.Format("「{0}」值由『{1}』改為『{2}』。", e.FieldName, e.OldValue, e.NewValue);
                List<string> Fields = new List<string>() { e.FieldName };
                ErrorType ErrorType = ErrorType.Correct;
                ValidatorType ValidatorType = ValidatorType.Field;

                MessageItem item = new MessageItem(ErrorType, ValidatorType, Description, Fields);

                RowMessages[e.Row.Position].MessageItems.Add(item);

                AutoCorrectCount++;
            }
            else //就不處理
            {

            }
            #endregion
        }

        /// <summary>
        /// 錯誤或警告事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void docValidate_ErrorCaptured(object sender, ErrorCapturedEventArgs e)
        {
            string Description = e.Description;
            List<string> Fields = new List<string>() { e.FieldName };
            ErrorType ErrorType = (e.ErrorType == EMBA.DocumentValidator.ErrorType.Error) ? ErrorType.Error : ErrorType.Warning;
            ValidatorType ValidationType = ValidatorType.Field;

            MessageItem item = new MessageItem(ErrorType, ValidationType, Description, Fields);

            RowMessages[e.Row.Position].MessageItems.Add(item);

            if (e.ErrorType == EMBA.DocumentValidator.ErrorType.Error)
                ErrorCount++;
            else
                WarningCount++;
        }
    }
}