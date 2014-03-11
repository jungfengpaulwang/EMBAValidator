using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aspose.Cells;
using FISCA.Presentation.Controls;

namespace EMBA.Validator
{
    /// <summary>
    /// 協助產生驗證報告類別
    /// </summary>
    public class OutputBuilder
    {
        private ValidatePair mCurrentPair;
        private List<ValidatePair> mValidatorPairs;
        private string mOutputFile;
        private SheetHelper mSheet;
        private ReportProgress mReportProgress;

        /// <summary>
        /// 判斷來源檔案路徑是否都屬於同一個檔案，並且輸出選項只為Full
        /// </summary>
        /// <param name="ValidatorPairs"></param>
        /// <param name="OutputOptions"></param>
        /// <returns></returns>
        static public bool SameSource(IEnumerable<ValidatePair> ValidatorPairs,OutputOptions OutputOptions)
        {            
            return ValidatorPairs.Select(x => x.DataFile).Distinct().Count() == 1
                && OutputOptions == OutputOptions.Full;
        }

        /// <summary>
        /// 建構式，將Validator的Validate方法參數傳入
        /// </summary>
        /// <param name="ValidatorPairs"></param>
        /// <param name="OutputFile"></param>
        /// <param name="OutputOptions"></param>
        public OutputBuilder(IEnumerable<ValidatePair> ValidatorPairs,
            string OutputFile,
            OutputOptions OutputOptions,
            ReportProgress ReportProgress)
        {
            //初始化從Validator傳來的參數
            mValidatorPairs = ValidatorPairs.ToList();
            mOutputFile = OutputFile;
            //預設讀取第一個Excel
            mSheet = new SheetHelper(mValidatorPairs[0].DataFile, mValidatorPairs[0].DataSheet);
            mCurrentPair = mValidatorPairs[0];
            mReportProgress = ReportProgress;
        }

        ///// <summary>
        ///// 是否為最佳化模式，亦即ValidatePair中的DataFile都屬於同個檔案。
        ///// </summary>
        //public bool IsOptimizeMode 
        //{ 
        //    get 
        //    { 
        //        return mIsOptimizeMode;
        //    }
        //}

        /// <summary>
        /// 取得作用中的活頁簿
        /// </summary>
        public SheetHelper Sheet { get { return mSheet; } }

        /// <summary>
        /// 最後已有『驗證訊息』欄位，則將其下所有欄位值清空，若無的話加上『驗證訊息』表頭
        /// </summary>
        public void InitialMessageHeader()
        {
            int ValidatorMessageColumnIndex;
            
            string MaxDataColumnValue = "" + Sheet.Sheet.Cells[0, Sheet.Sheet.Cells.MaxDataColumn].Value;
            
            if (MaxDataColumnValue.Contains("驗證訊息")) //如果已經有此欄位
            {
                ValidatorMessageColumnIndex = Sheet.Sheet.Cells.MaxDataColumn;
                for (int x = 1; x <= Sheet.Sheet.Cells.MaxDataRow; x++) //清空此欄位的內容
                    Sheet.Sheet.Cells[x, ValidatorMessageColumnIndex].PutValue("");
            }
            else //如果沒有此欄位
            {
                ValidatorMessageColumnIndex = Sheet.Sheet.Cells.MaxDataColumn + 1;
            }

            Sheet.Sheet.Cells[0, ValidatorMessageColumnIndex].PutValue("驗證訊息");
        }

        /// <summary>
        /// 設定欄位驗證訊息
        /// </summary>
        /// <param name="Message"></param>
        public void SetHeaderMessage(string Message)
        {
            Sheet.Sheet.Cells[0, Sheet.Sheet.Cells.MaxDataColumn].PutValue("驗證訊息" + (!string.IsNullOrEmpty(Message) ? "\r\n" + Message : ""));
            Sheet.Sheet.Cells[0, Sheet.Sheet.Cells.MaxDataColumn].Style.IsTextWrapped = true;
        }

        /// <summary>
        /// 寫入單筆資料驗證訊息
        /// </summary>
        /// <param name="Position"></param>
        /// <param name="Message"></param>
        public void SetMessage(int Position,string Message)
        {
            Cell cell = Sheet.Sheet.Cells[Position, Sheet.Sheet.Cells.MaxDataColumn];
            cell.PutValue(Message);
            cell.Style.IsTextWrapped = true;
        }

        /// <summary>
        /// 寫入多筆資料驗證訊息
        /// </summary>
        /// <param name="Messages"></param>
        public void SetMessages(RowMessages Messages)
        {
            if (mReportProgress!=null)
                mReportProgress(mCurrentPair, 0, 0, 0, "寫入驗證訊息:", 0);

            int WriteMessageProgress = 0;
            int WriteMessageTotal = Messages.Count;

            Dictionary<int, string> BestMessageDescription = new Dictionary<int, string>();

            foreach (RowMessage Message in Messages)
            {
                List<MessageItem> Items = Message.BestMessageItems;
                BestMessageDescription.Add(Message.Position, Items.GetDescription());
            }

            foreach (RowMessage Message in Messages)
            {
                string strMessage = BestMessageDescription[Message.Position];

                SetMessage(Message.Position, strMessage);

                if (mReportProgress!=null)
                    mReportProgress(mCurrentPair, Messages.ErrorCount, Messages.WarningCount, Messages.AutoCorrectCount, "寫入驗證訊息: " + mCurrentPair.DataSheet, (int)(++WriteMessageProgress * 100 / WriteMessageTotal));
            }

            //更新正確的錯誤、警告及自動修正數目
            Messages.UpdateErrorWarningAndAutoCorrectCount();

            if (mReportProgress!=null)
                mReportProgress(mCurrentPair, Messages.ErrorCount, Messages.WarningCount, Messages.AutoCorrectCount, "寫入驗證訊息:", 100);
        }

        /// <summary>
        /// 依『驗證訊息』欄位中的內容調整最適欄寬
        /// </summary>
        public void AutoFitMessage()
        {
            if (mReportProgress!=null)
                mReportProgress(mCurrentPair, 0,0,0, "調整最適欄寬: " + mCurrentPair.DataSheet, 0);
            
            Sheet.Sheet.AutoFitColumn(Sheet.Sheet.Cells.MaxDataColumn);

            Sheet.Sheet.AutoFitRows();

            if (mReportProgress!=null)
                mReportProgress(mCurrentPair, 0, 0, 0, "調整最適欄寬: " + mCurrentPair.DataSheet, 100);
        }

        /// <summary>
        /// 根據ValidatePair來切換作用中的活頁簿
        /// 1.若是最佳化模式，則直接切換資料表。
        /// 2.若非最佳化模式，則開啟新的活頁簿。
        /// </summary>
        /// <param name="Pair"></param>
        public void Switch(ValidatePair Pair)
        {
            //若是最佳化模式或是與前次驗證的檔案路徑相同，則直接切換資料表即可
            //if (IsOptimizeMode || Pair.DataFile.Equals(mCurrentPair.DataFile))
                mSheet.SwitchSeet(Pair.DataSheet);
            //else
            //    mSheet = new SheetHelper(Pair.DataFile, Pair.DataSheet);

            mCurrentPair = Pair;
        }

        //不支援OutputOptions模式
        ///// <summary>
        ///// 驗證完後產生報告，其中參數為有驗證訊息的列編號
        ///// </summary>
        ///// <param name="Positions"></param>
        //public void ProduceReport(List<int> Positions)
        //{
        //    //若為最佳化模式，則最後直接將來源檔案存檔即可，不需要再複製到另個活頁簿。
        //    if (IsOptimizeMode)
        //        return;

        //    bool full = (mOutputOptions & OutputOptions.Full) == OutputOptions.Full;
        //    bool correct = (mOutputOptions & OutputOptions.Correct) == OutputOptions.Correct;
        //    bool error = (mOutputOptions & OutputOptions.Error) == OutputOptions.Error;

        //    Worksheet fullSheet = null;
        //    Worksheet correctSheet = null;
        //    Worksheet errorSheet = null;

        //    if (full)
        //        fullSheet = mOutputBook.Worksheets[mOutputBook.Worksheets.Add()];
        //    if (correct)
        //    {
        //        correctSheet = mOutputBook.Worksheets[mOutputBook.Worksheets.Add()];
        //        correctSheet.Cells.CopyRow(Sheet.Sheet.Cells, 0, 0);
        //    }
        //    if (error)
        //    {
        //        errorSheet = mOutputBook.Worksheets[mOutputBook.Worksheets.Add()];
        //        errorSheet.Cells.CopyRow(Sheet.Sheet.Cells, 0, 0);
        //    }

        //    int correctCount = 1;
        //    int errorCount = 1;
        //    int DataRowCount = Sheet.DataRowCount;

        //    for (int i = Sheet.FirstDataRowIndex; i <= Sheet.DataRowCount; i++)
        //    {
        //        if (!Positions.Contains(i) && correct)
        //            correctSheet.Cells.CopyRow(Sheet.Sheet.Cells, i, correctCount++);
        //        else if (Positions.Contains(i) && error)
        //            errorSheet.Cells.CopyRow(Sheet.Sheet.Cells, i, errorCount++);

        //        Validator.ReportProgress(mCurrentPair, 0, 0, 0, "產生驗證報告: " + mCurrentPair.DataSheet, (int)(i * 100 / DataRowCount));

        //        //if (Progress != null)
        //        //{
        //        //    ValidatingPair obj = new ValidatingPair();

        //        //    obj.ErrorCount = RowMessages.ErrorCount;
        //        //    obj.WarningCount = RowMessages.WarningCount;
        //        //    obj.AutoCorrectCount = RowMessages.AutoCorrectCount;

        //        //    obj.Message = "產生驗證報告: " + Sheet.Sheet.Name;
        //        //    obj.DataFile = Path.GetFileName(each.DataFile);
        //        //    obj.DataSheet = each.DataSheet;

        //        //    Progress(obj, (int)(++count * 100 / total));
        //        //}
        //    }

        //    if (full)
        //    {
        //        fullSheet.Copy(Sheet.Sheet);
        //        fullSheet.AutoFitColumn(Sheet.Sheet.Cells.MaxDataColumn);
        //        fullSheet.AutoFitRows();
        //        fullSheet.Name = Sheet.Sheet.Name;
        //    }
        //    if (correct)
        //    {
        //        correctSheet.AutoFitColumn(Sheet.Sheet.Cells.MaxDataColumn);
        //        correctSheet.AutoFitRows();
        //        correctSheet.Name = Sheet.Sheet.Name + "(正確)";
        //    }
        //    if (error)
        //    {
        //        errorSheet.AutoFitColumn(Sheet.Sheet.Cells.MaxDataColumn);
        //        errorSheet.AutoFitRows();
        //        errorSheet.Name = Sheet.Sheet.Name + "(錯誤)";
        //    }
        //}

        /// <summary>
        /// 將驗證報告存檔
        /// </summary>
        public void Save()
        {
            try
            {
                if (mReportProgress!=null)
                    mReportProgress(null, 0, 0, 0, "驗證報告存檔中!", 0);
                
                //if (IsOptimizeMode)
                    Sheet.Book.Save(mOutputFile);
                //else
                //    mOutputBook.Save(mOutputFile);

                if (mReportProgress!=null)
                    mReportProgress(null, 0, 0, 0, "驗證報告存檔完成!", 100);
            }
            catch (System.OutOfMemoryException ex)
            {
                StringBuilder strBuilder = new StringBuilder();

                strBuilder.AppendLine("『" + mOutputFile + "』檔案在儲存時『記憶體不足（Out Of Memory）』!!");
                strBuilder.AppendLine("解決方案一：將系統未使用程式關閉增加可用之記憶體資源。");
                strBuilder.AppendLine("解決方案二：分批匯入資料。");

                MsgBox.Show(strBuilder.ToString());

                //SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
            }
            catch (System.IO.IOException ex)
            {
                MsgBox.Show("檔案可能開啟中!\n" + ex.Message);
                //SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
            }
            catch (Exception ex)
            {
                MsgBox.Show("驗證報告儲存發生錯誤!\n" + ex.Message);
                //SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
            }
        }
    }
}