using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Validator
{
    public class RowMessage
    {
        public RowMessage(int Position)
        {
            MessageItems = new List<MessageItem>();
            this.Position = Position;
        }

        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
        public int AutoCorrectCount { get; set; }

        public List<MessageItem> MessageItems { get; set; }

        public List<MessageItem> BestMessageItems
        {
            get
            {
                List<MessageItem> resultList = new List<MessageItem>();
                List<MessageItem> fieldList = new List<MessageItem>();

                foreach (MessageItem item in MessageItems)
                {
                    if (item.ValidatorType == ValidatorType.Row)
                        resultList.Add(item);
                    else
                        fieldList.Add(item);
                }

                int index = 0;
                foreach (var thisGroup in fieldList.GroupBy(x => x.FieldsString))
                {
                    List<MessageItem> Items = new List<MessageItem>();
                    foreach (var Value in thisGroup)
                        Items.Add(Value);
                    resultList.Insert(index++, Items.GetBest());
                }

                //這樣寫可以嗎？
                //fieldList.GroupBy(x => x.FieldsString).ToList().ForEach((thisGroup) =>
                //{
                //    resultList.Insert(index++, thisGroup.ToList().GetBest());
                //});

                foreach (var errorTypeGroup in resultList.GroupBy(x => x.ErrorType))
                {
                    if (errorTypeGroup.Key == ErrorType.Correct)
                        AutoCorrectCount = errorTypeGroup.Count();
                    else if (errorTypeGroup.Key == ErrorType.Error)
                        ErrorCount = errorTypeGroup.Count();
                    else if (errorTypeGroup.Key == ErrorType.Warning)
                        WarningCount = errorTypeGroup.Count();
                }

                return resultList;
            }
        }

        /// <summary>
        /// RowMessage 對應的資料行位置。
        /// </summary>
        public int Position { get; private set; }
    }
}
