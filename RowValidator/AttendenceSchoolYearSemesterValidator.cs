using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EMBA.DocumentValidator;

namespace EMBA.Validator
{
    public class AttendenceSchoolYearSemesterValidator : EMBA.DocumentValidator.IRowVaildator
    {
        //key: 學號+日期
        //value: 學年度+學期
        private Dictionary<string, string> StudentDate { get; set; }

        public AttendenceSchoolYearSemesterValidator()
        {
            StudentDate = new Dictionary<string, string>();
        }

        #region IRowVaildator 成員

        public bool Validate(IRowStream Value)
        {
            string studentNumber = Value.GetValue("學號");
            string date = Value.GetValue("日期");
            string schoolYear = Value.GetValue("學年度");
            string semester = Value.GetValue("學期");

            string key = Combine(studentNumber, date);
            string sems = Combine(schoolYear, semester);

            if (!StudentDate.ContainsKey(key))
                StudentDate.Add(key, sems);
            if (sems != StudentDate[key]) return false;
            else return true;
        }

        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        #endregion

        private string Combine(string a, string b)
        {
            return a + "_" + b;
        }

        //private class Sems
        //{
        //    private Dictionary<string, List<int>> SemsRowIndexes { get; set; }

        //    public Sems()
        //    {
        //        SemsRowIndexes = new Dictionary<string, List<int>>();
        //    }

        //    public List<int> Add(string sems, int rowIndex)
        //    {
        //        if (!SemsRowIndexes.ContainsKey(sems))
        //            SemsRowIndexes.Add(sems, new List<int>());
        //        SemsRowIndexes[sems].Add(rowIndex);

        //        if (SemsRowIndexes.Count > 1)
        //            return GetRowIndexes();
        //    }

        //    private List<int> GetRowIndexes()
        //    {
        //        List<int> list = new List<int>();
        //        SemsRowIndexes.Values.ToList().ForEach(x => list.AddRange(x));
        //        return list;
        //    }
        //}
    }
}
