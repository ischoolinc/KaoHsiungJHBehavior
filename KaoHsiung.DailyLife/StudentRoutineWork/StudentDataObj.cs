using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JHSchool.Data;
using System.Xml;
using FISCA.DSAUtil;
using JHSchool.Behavior.BusinessLogic;
using K12.Data;

namespace KaoHsiung.DailyLife.StudentRoutineWork
{
    class StudentDataObj
    {
        #region 物件資料欄位

        /// <summary>
        /// 學生基本資料物件
        /// </summary>
        public JHStudentRecord StudentRecord { get; set; }

        /// <summary>
        /// 戶籍電話
        /// </summary>
        public string PhonePermanent { get; set; }

        /// <summary>
        /// 聯絡電話
        /// </summary>
        public string PhoneContact { get; set; }

        /// <summary>
        /// 監護人姓名
        /// </summary>
        public string CustodianName { get; set; }

        /// <summary>
        /// 戶籍地址
        /// </summary>
        public string AddressPermanent { get; set; }

        /// <summary>
        /// 聯絡地址
        /// </summary>
        public string AddressMailing { get; set; }

        /// <summary>
        /// 學期歷程
        /// </summary>
        public JHSemesterHistoryRecord SemesterHistory { get; set; }

        /// <summary>
        /// 獎勵資料清單
        /// </summary>
        public List<MeritRecord> ListMerit = new List<MeritRecord>();

        /// <summary>
        /// 懲戒資料清單
        /// </summary>
        public List<DemeritRecord> ListDeMerit = new List<DemeritRecord>();

        /// <summary>
        /// 日常生活表現
        /// </summary>
        public List<JHMoralScoreRecord> ListMoralScore = new List<JHMoralScoreRecord>();

        /// <summary>
        /// 異動記錄
        /// </summary>
        public List<JHUpdateRecordRecord> ListUpdateRecord = new List<JHUpdateRecordRecord>();

        /// <summary>
        /// 社團活動記錄
        /// </summary>
        public List<AssnCode> ListAssnCode = new List<AssnCode>();

        /// <summary>
        /// 自動統計缺曠獎懲
        /// </summary>
        public List<AutoSummaryRecord> ListAutoSummary = new List<AutoSummaryRecord>();

        /// <summary>
        /// 日常生活表現,學年度學期,標題:內容
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> TextScoreDic = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 統計資料,學年度學期,假別:內容
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> SummaryDic = new Dictionary<string, Dictionary<string, string>>();

        /// <summary>
        /// 社團活動
        /// </summary>
        //public Dictionary<string, string> AssnDic = new Dictionary<string, string>();

        /// <summary>
        /// 努力程度物件
        /// </summary>
        private EffortMapper Effor = new EffortMapper();

        /// <summary>
        /// 服務學習時數
        /// </summary>
        public List<SLRecord> ListSLR = new List<SLRecord>();

        #endregion

        /// <summary>
        /// 初始化資料內容
        /// </summary>
        public void SetupData()
        {
            //取得新生資料
            NewStudent();

            //資料排序
            DataSort();

            //處理學期歷程
            ExecuteSemesterSet();

            //處理社團活動
            //ExecuteAssnCode();

            //處理日常生活表現
            ExecuteMoralScore();

        }

        #region 新生異動

        /// <summary>
        /// 畢業國小
        /// </summary>
        public string UpdataGraduateSchool { get; set; }

        /// <summary>
        /// 入學核准日期
        /// </summary>
        public string UpdataADDate { get; set; }

        /// <summary>
        /// 入學核准文號
        /// </summary>
        public string UpdataADNumber { get; set; }

        /// <summary>
        /// 取得新生入學資訊
        /// </summary>
        private void NewStudent()
        {
            foreach (JHUpdateRecordRecord each in ListUpdateRecord)
            {
                if (each.UpdateCode == "1")
                {
                    UpdataGraduateSchool = each.GraduateSchool;
                    UpdataADDate = each.ADDate;
                    UpdataADNumber = each.ADNumber;
                }
            }
        }

        #endregion

        #region 排序

        private void DataSort()
        {
            ListMerit.Sort(new Comparison<MeritRecord>(MeritSort));
            ListDeMerit.Sort(new Comparison<DemeritRecord>(DeMeritSort));
            ListUpdateRecord.Sort(new Comparison<JHUpdateRecordRecord>(UpdateRecordSort));
            ListMoralScore.Sort(new Comparison<JHMoralScoreRecord>(MoralScoreSort));

        }

        private int MeritSort(MeritRecord x, MeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        private int DeMeritSort(DemeritRecord x, DemeritRecord y)
        {
            return x.OccurDate.CompareTo(y.OccurDate);
        }

        private int UpdateRecordSort(JHUpdateRecordRecord x, JHUpdateRecordRecord y)
        {
            DateTime dt1;
            DateTime.TryParse(x.UpdateDate, out dt1);
            DateTime dt2;
            DateTime.TryParse(y.UpdateDate, out dt2);
            return dt1.CompareTo(dt2);
        }

        private int MoralScoreSort(JHMoralScoreRecord x, JHMoralScoreRecord y)
        {
            string xx = x.SchoolYear.ToString() + x.Semester.ToString();
            string yy = y.SchoolYear.ToString() + y.Semester.ToString();
            return xx.CompareTo(yy);
        }

        #endregion

        #region 學期歷程

        //班級座號
        public string GradeYear11 { get; set; }
        public string GradeYear12 { get; set; }
        public string GradeYear21 { get; set; }
        public string GradeYear22 { get; set; }
        public string GradeYear31 { get; set; }
        public string GradeYear32 { get; set; }

        //應到日數
        public Dictionary<string, string> SchoolDay = new Dictionary<string, string>();

        private int SortSemester(K12.Data.SemesterHistoryItem t1, K12.Data.SemesterHistoryItem t2)
        {
            string schoolyear1 = t1.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear1 += t1.Semester.ToString().PadLeft(5, '0');

            string schoolyear2 = t2.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear2 += t2.Semester.ToString().PadLeft(5, '0');

            return schoolyear1.CompareTo(schoolyear2);
        }

        private int SortMoralScore(JHMoralScoreRecord x1, JHMoralScoreRecord x2)
        {
            string schoolyear1 = x1.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear1 += x1.Semester.ToString().PadLeft(5, '0');

            string schoolyear2 = x2.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear2 += x2.Semester.ToString().PadLeft(5, '0');

            return schoolyear1.CompareTo(schoolyear2);

        }

        private int SortAutoSummary(AutoSummaryRecord x1, AutoSummaryRecord x2)
        {
            string schoolyear1 = x1.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear1 += x1.Semester.ToString().PadLeft(5, '0');

            string schoolyear2 = x2.SchoolYear.ToString().PadLeft(5, '0');
            schoolyear2 += x2.Semester.ToString().PadLeft(5, '0');

            return schoolyear1.CompareTo(schoolyear2);

        }
        //處理學期歷程資料
        private void ExecuteSemesterSet()
        {
            List<K12.Data.SemesterHistoryItem> list = SemesterHistory.SemesterHistoryItems;
            list.Sort(SortSemester);
            foreach (K12.Data.SemesterHistoryItem each in list)
            {
                string SchoolYearSemester = each.SchoolYear.ToString() + "/" + each.Semester.ToString();

                if (SchoolDay.ContainsKey(SchoolYearSemester)) //如果重覆則跳出
                    continue;

                string Day = each.SchoolDayCount.HasValue ? each.SchoolDayCount.Value.ToString() : "";

                SchoolDay.Add(SchoolYearSemester, Day);

                string ClassNameAndSeatNo = each.ClassName + "/" + (each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "");
                if (each.GradeYear == 1 || each.GradeYear == 7)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear11 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear12 = ClassNameAndSeatNo;
                    }
                }
                else if (each.GradeYear == 2 || each.GradeYear == 8)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear21 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear22 = ClassNameAndSeatNo;
                    }
                }
                else if (each.GradeYear == 3 || each.GradeYear == 9)
                {
                    if (each.Semester == 1)
                    {
                        GradeYear31 = ClassNameAndSeatNo;
                    }
                    else
                    {
                        GradeYear32 = ClassNameAndSeatNo;
                    }
                }
            }
        }

        #endregion

        /// <summary>
        /// 社團活動資料整理
        /// </summary>
        //private void ExecuteAssnCode()
        //{
        //    foreach (AssnCode each in ListAssnCode)
        //    {
        //        string AssnScore = "";

        //        if (each.Scores == "")
        //            continue;

        //        string SchoolYearSemester = each.SchoolYear + "/" + each.Semester;

        //        XmlElement xml = DSXmlHelper.LoadXml(each.Scores);
        //        foreach (XmlElement item in xml.SelectNodes("Item"))
        //        {
        //            string x1 = "社團表現：" + item.GetAttribute("Score");
        //            string x2 = "社團表現：" + item.GetAttribute("Effort");
        //            string x3 = "社團表現：" + item.GetAttribute("Text");

        //            if (x1 != "社團表現：")
        //            {
        //                AssnScore += "" + x1 + "。";
        //            }

        //            if (x2 != "社團表現：")
        //            {
        //                int jj;
        //                if (int.TryParse(item.GetAttribute("Effort"), out jj))
        //                {
        //                    AssnScore += "" + Effor.GetTextByCode(jj) + "。";
        //                }
        //                else
        //                {
        //                    AssnScore += "" + x2 + "。";
        //                }
        //            }

        //            if (x3 != "社團表現：")
        //            {
        //                AssnScore += "" + x3 + "。";
        //            }
        //        }
        //        if (AssnScore != string.Empty)
        //        {
        //            //2013/7/2 - 社團成績只取1筆資料
        //            if (!AssnDic.ContainsKey(SchoolYearSemester))
        //            {
        //                AssnDic.Add(SchoolYearSemester, AssnScore);
        //            }
        //        }
        //    }
        //}

        /// <summary>
        /// 日常生活表現資料整理
        /// </summary>
        private void ExecuteMoralScore()
        {
            ListAutoSummary.Sort(SortAutoSummary);

            foreach (AutoSummaryRecord auto in ListAutoSummary)
            {
                //學年度學期
                string SchoolYearSemester = auto.SchoolYear.ToString() + "/" + auto.Semester.ToString();

                if (!SchoolDay.ContainsKey(SchoolYearSemester))
                    continue;

                #region 處理出缺席統計

                if (!SummaryDic.ContainsKey(SchoolYearSemester))
                {
                    SummaryDic.Add(SchoolYearSemester, new Dictionary<string, string>());
                }

                AutoSummaryRecord Asr = auto;

                if (Asr != null)
                {

                    DSXmlHelper helper2 = new DSXmlHelper(Asr.AutoSummary);

                    if (helper2.GetElement("AttendanceStatistics") != null)
                    {
                        foreach (XmlElement item in helper2.GetElements("AttendanceStatistics/Absence"))
                        {

                            SummaryDic[SchoolYearSemester].Add(item.GetAttribute("Name") + item.GetAttribute("PeriodType"), item.GetAttribute("Count"));
                        }
                    }

                    if (helper2.GetElement("DisciplineStatistics/Merit") != null)
                    {
                        SummaryDic[SchoolYearSemester].Add("大功", helper2.GetAttribute("DisciplineStatistics/Merit/@A"));
                        SummaryDic[SchoolYearSemester].Add("小功", helper2.GetAttribute("DisciplineStatistics/Merit/@B"));
                        SummaryDic[SchoolYearSemester].Add("嘉獎", helper2.GetAttribute("DisciplineStatistics/Merit/@C"));
                    }

                    if (helper2.GetElement("DisciplineStatistics/Demerit") != null)
                    {
                        SummaryDic[SchoolYearSemester].Add("大過", helper2.GetAttribute("DisciplineStatistics/Demerit/@A"));
                        SummaryDic[SchoolYearSemester].Add("小過", helper2.GetAttribute("DisciplineStatistics/Demerit/@B"));
                        SummaryDic[SchoolYearSemester].Add("警告", helper2.GetAttribute("DisciplineStatistics/Demerit/@C"));
                    }
                }
                #endregion
            }

            ListMoralScore.Sort(SortMoralScore);

            foreach (JHMoralScoreRecord each in ListMoralScore)
            {
                //學年度學期
                string SchoolYearSemester = each.SchoolYear.ToString() + "/" + each.Semester.ToString();

                if (!SchoolDay.ContainsKey(SchoolYearSemester))
                    continue;

                #region 日常生活表現
                if (!TextScoreDic.ContainsKey(SchoolYearSemester))
                {
                    TextScoreDic.Add(SchoolYearSemester, new Dictionary<string, string>());
                }

                DSXmlHelper helper1 = new DSXmlHelper(each.TextScore);

                if (helper1.GetElements("DailyBehavior/Item").Length != 0)
                {
                    foreach (XmlElement item in helper1.GetElements("DailyBehavior/Item"))
                    {
                        TextScoreDic[SchoolYearSemester].Add(item.GetAttribute("Name"), item.GetAttribute("Degree"));
                    }
                }

                if (helper1.GetElements("GroupActivity/Item").Length != 0)
                {
                    //string name = helper1.GetElement("GroupActivity").GetAttribute("Name");
                    //if (name != "")
                    //{
                    string GroupActivityText = string.Empty;

                    //社團處理 - 2022/1/18 高雄小組決議移除
                    //if (AssnDic.ContainsKey(SchoolYearSemester))
                    //{
                    //    GroupActivityText += AssnDic[SchoolYearSemester];
                    //}

                    foreach (XmlElement item in helper1.GetElements("GroupActivity/Item"))
                    {

                        if (item.GetAttribute("Name") == "社團活動" || item.GetAttribute("Name") == "自治活動")
                            continue;

                        if (!string.IsNullOrEmpty(item.GetAttribute("Degree")))
                        {
                            GroupActivityText += item.GetAttribute("Name") + "：" + item.GetAttribute("Degree") + "。";
                        }
                        if (!string.IsNullOrEmpty(item.GetAttribute("Description")))
                        {
                            GroupActivityText += item.GetAttribute("Name") + "：" + item.GetAttribute("Description") + "。";
                        }
                    }
                    TextScoreDic[SchoolYearSemester].Add("GroupActivity", GroupActivityText);
                    //}

                }

                if (helper1.GetElements("PublicService/Item").Length != 0)
                {
                    //string name = helper1.GetElement("PublicService").GetAttribute("Name");
                    //if (name != "")
                    //{
                    string PublicServiceText = string.Empty;
                    foreach (XmlElement item in helper1.GetElements("PublicService/Item"))
                    {
                        if (!string.IsNullOrEmpty(item.GetAttribute("Description")))
                        {
                            PublicServiceText += item.GetAttribute("Name") + "：" + item.GetAttribute("Description") + "。";
                        }
                    }
                    TextScoreDic[SchoolYearSemester].Add("PublicService", PublicServiceText);
                    //}
                }

                if (helper1.GetElements("SchoolSpecial/Item").Length != 0)
                {
                    //string name = helper1.GetElement("SchoolSpecial").GetAttribute("Name");
                    //if (name != "")
                    //{
                    string SchoolSpecialText = string.Empty;
                    foreach (XmlElement item in helper1.GetElements("SchoolSpecial/Item"))
                    {
                        if (!string.IsNullOrEmpty(item.GetAttribute("Description")))
                        {
                            SchoolSpecialText += item.GetAttribute("Name") + "：" + item.GetAttribute("Description");
                        }
                    }
                    TextScoreDic[SchoolYearSemester].Add("SchoolSpecial", SchoolSpecialText);
                    //}
                }


                if (helper1.GetElement("DailyLifeRecommend") != null)
                {
                    //string name = helper1.GetElement("DailyLifeRecommend").GetAttribute("Name");
                    //if (name != "")
                    //{
                    string Description = helper1.GetElement("DailyLifeRecommend").GetAttribute("Description");
                    TextScoreDic[SchoolYearSemester].Add("DailyLifeRecommend", Description);
                    //}
                }
                #endregion

            }
        }
    }
}
