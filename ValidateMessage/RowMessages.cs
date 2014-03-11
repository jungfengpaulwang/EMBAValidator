using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace EMBA.Validator
{
    public class RowMessages : IEnumerable<RowMessage>
    {
        private Dictionary<int, RowMessage> Messages { get; set; }

        public RowMessages()
        {
            Messages = new Dictionary<int, RowMessage>();
        }

        public int Count { get { return Messages.Count; } }

        public int ErrorCount { get; private set; }

        public int WarningCount { get; private set; }

        public int AutoCorrectCount { get; private set; }

        public void UpdateErrorWarningAndAutoCorrectCount()
        {
            ErrorCount = 0;

            foreach (RowMessage Row in Messages.Values)
            {
                //MessageBox.Show("" + Row.ErrorCount);
                ErrorCount +=  Row.ErrorCount;
            }

            WarningCount = Messages.Values.Select(x => x.WarningCount).Sum();
            AutoCorrectCount = Messages.Values.Select(x => x.AutoCorrectCount).Sum();
        }

        public List<int> Positions
        {
            get
            {
                return Messages.Keys.ToList();
            }
        }

        public RowMessage this[int Position]
        {
            get
            {
                if (!Messages.ContainsKey(Position))
                    Messages.Add(Position, new RowMessage(Position));

                return Messages[Position];

            }
            set
            {
                if (!Messages.ContainsKey(Position))
                    Messages.Add(Position, new RowMessage(Position));

                if (value.Position != Position)
                    throw new Exception();

                Messages[Position] = value;
            }
        }

        /// <summary>
        /// 排序 RowMessages。
        /// 依 Position，由小到大。
        /// </summary>
        public void Sort()
        {
            //Messages.Sort(delegate(RowMessage x, RowMessage y)
            //{
            //    return x.Position.CompareTo(y.Position);
            //});
        }

        public void Clear()
        {
            Messages.Clear();
        }

        #region IEnumerable<RowMessage> 成員

        public IEnumerator<RowMessage> GetEnumerator()
        {
            return Messages.Values.GetEnumerator();
        }

        #endregion

        #region IEnumerable 成員

        IEnumerator IEnumerable.GetEnumerator()
        {
            return Messages.Values.GetEnumerator();
        }

        #endregion
    }
}
