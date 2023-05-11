using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using System.Xml;
using JHSchool.Behavior;
using FISCA.DSAUtil;
using K12.Data.Utility;

namespace KaoHsiung.DailyLife
{
    class ChangeToCDS
    {
        SummaryObj obj;


        /// <summary>
        /// 傳入相關資訊,進行計算(扣除明細)
        /// </summary>
        public ChangeToCDS(string SaveRefStudent, int SaveSchoolYear, int SaveSemester,XmlElement xml)
        {
            List<JHDemeritRecord> DemeritList = JHDemerit.SelectBySchoolYearAndSemester(SaveRefStudent, SaveSchoolYear, SaveSemester);

            List<JHMeritRecord> MeritList = JHMerit.SelectBySchoolYearAndSemester(SaveRefStudent, SaveSchoolYear, SaveSemester);

            List<JHAttendanceRecord> AttendanceList = JHAttendance.SelectBySchoolYearAndSemester
                (new JHStudentRecord[] { JHStudent.SelectByID(SaveRefStudent) }, SaveSchoolYear, SaveSemester);

            GetPeriodTypeItems();

            obj = new SummaryObj(xml);

            //扣掉獎勵
            foreach (JHMeritRecord each in MeritList)
            {
                if (each.MeritA.HasValue)
                {
                    obj.MeritA = obj.MeritA - each.MeritA.Value;
                }
                if (each.MeritB.HasValue)
                {
                    obj.MeritB = obj.MeritB - each.MeritB.Value;
                }
                if (each.MeritC.HasValue)
                {
                    obj.MeritC = obj.MeritC - each.MeritC.Value;
                }
            }

            //扣掉懲戒
            foreach (JHDemeritRecord each in DemeritList)
            {
                //20230511 - 如果銷過就跳過
                if (each.Cleared == "是")
                    continue;

                if (each.DemeritA.HasValue)
                {
                    obj.DemeritA = obj.DemeritA - each.DemeritA.Value;
                }
                if (each.DemeritB.HasValue)
                {
                    obj.DemeritB = obj.DemeritB - each.DemeritB.Value;
                }
                if (each.DemeritC.HasValue)
                {
                    obj.DemeritC = obj.DemeritC - each.DemeritC.Value;
                }
            }

            Dictionary<string, int> periodList = new Dictionary<string, int>();

            //扣掉缺曠
            foreach (JHAttendanceRecord attend in AttendanceList)
            {
                foreach (K12.Data.AttendancePeriod period in attend.PeriodDetail)
                {
                    if (PeriodTypeDic.ContainsKey(period.Period))
                    {
                        foreach (AttendanceObj each in obj.AttendanceList)
                        {
                            if (each.PeriodType == PeriodTypeDic[period.Period] && each.Name == period.AbsenceType)
                            {
                                each.Count--;
                            }
                        }
                    }
                }
            }

        }

        public XmlElement GetXmlElement()
        {
            return obj.GetAllXmlElement();
        }

        Dictionary<string, string> PeriodTypeDic = new Dictionary<string, string>();

        public List<string> GetPeriodTypeItems()
        {
            #region 取得節次類型

            PeriodTypeDic.Clear();

            string targetService = "SmartSchool.Config.GetList";

            List<string> list = new List<string>();

            DSXmlHelper helper = new DSXmlHelper("GetListRequest");
            helper.AddElement("Field");
            helper.AddElement("Field", "Content", "");
            helper.AddElement("Condition");
            helper.AddElement("Condition", "Name", "節次對照表");

            DSRequest req = new DSRequest(helper.BaseElement);
            DSResponse rsp = DSAServices.CallService(targetService, req);

            foreach (XmlElement element in rsp.GetContent().GetElements("List/Periods/Period"))
            {
                string type = element.GetAttribute("Type");

                if (!list.Contains(type))
                {
                    list.Add(type);
                }

                PeriodTypeDic.Add(element.GetAttribute("Name"), element.GetAttribute("Type"));
            }
            return list;
            #endregion
        }


    }
}
