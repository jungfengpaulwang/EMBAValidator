using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EMBA.Validator
{
    public class MessageItem
    {
        public MessageItem(ErrorType ErrorType, ValidatorType ValidatorType, string Message)
        {
            this.ErrorType = ErrorType;
            this.ValidatorType = ValidatorType;
            this.Message = Message;
            Fields = new List<string>();
        }

        public MessageItem(ErrorType ErrorType, ValidatorType ValidatorType, string Message, IEnumerable<string> RelatedFields)
            : this(ErrorType, ValidatorType, Message)
        {
            this.Fields.AddRange(RelatedFields);
        }

        public MessageItem(ErrorType ErrorType, ValidatorType ValidatorType, string Message, params string[] RelatedFields)
            : this(ErrorType, ValidatorType, Message)
        {
            this.Fields.AddRange(RelatedFields);
        }

        public ErrorType ErrorType { get; private set; }

        public ValidatorType ValidatorType { get; private set; }
        
        public List<string> Fields { get; private set; }

        public string FieldsString { get { return string.Join(",", Fields.ToArray()); } }
        
        public string Message { get; private set; }
    }

    public enum ErrorType
    {
        Correct, Error, Warning
    }

    public enum ValidatorType
    {
        Field, Row
    }
}
