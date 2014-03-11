using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Validator
{
    public static class MessageItemExtensionMethod
    {
        public static MessageItem GetBest(this List<MessageItem> Items)
        {
            foreach (MessageItem item in Items)
                if (item.ErrorType == ErrorType.Correct) return item;

            foreach (MessageItem item in Items)
                if (item.ErrorType == ErrorType.Error) return item;

            foreach (MessageItem item in Items)
                if (item.ErrorType == ErrorType.Warning) return item;

            return null;
        }

        public static string GetDescription(this List<MessageItem> Items)
        {
            return string.Join(
                "\n",
                Items.Select(
                    x => ((x.ErrorType == ErrorType.Error) ? "錯誤：" : (x.ErrorType == ErrorType.Warning ? "警告：" : "自動修正：")) + x.Message
                ).ToArray()
            );
        }
    }
}