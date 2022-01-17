using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool;
using JHSchool.Data;
using Aspose.Cells;
using System.IO;
using FISCA.DSAUtil;
using K12.Data.Utility;
using System.Xml;
using K12.Data.Configuration;
using Framework;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Evaluation.Calculation;
using KaoHsiung.DailyLife.日常生活表現總表.Calculation;

namespace KaoHsiung.DailyLife.日常生活表現總表
{
    public partial class ClassDailyLifeReport : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();
        BackgroundWorker BGW_new = new BackgroundWorker();
        private string SchoolYear;
        private string Semester;
        private bool IsPredic; //是否進行及格判斷
        private Dictionary<string, Dictionary<string, JHMoralScoreRecord>> SuperList;
        private Workbook wb = new Workbook();
        private Dictionary<string, int> ColumnInTitleIndex = new Dictionary<string, int>();
        EffortMapper Effor = new EffortMapper();
        GetMSRlist GetData;

        string jjMessage = "";

        private EvaluationResult er;
        List<string> StudentNotPassList = new List<string>();

        private int _sizeIndex = 0;

        public ClassDailyLifeReport()
        {
            InitializeComponent();
        }

        private void ClassDailyLifeReport_Load(object sender, EventArgs e)
        {
            #region 學年度學期
            string schoolYear = School.DefaultSchoolYear;
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 1).ToString());
            int x = cbSchoolYear.Items.Add((int.Parse(schoolYear)).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 1).ToString());
            cbSchoolYear.SelectedIndex = x;

            string semester = School.DefaultSemester;
            int z = cbSemester.Items.Add("1");
            int y = cbSemester.Items.Add("2");
            if (semester == "1")
            {
                cbSemester.SelectedIndex = z;
            }
            else
            {
                cbSemester.SelectedIndex = y;
            }
            #endregion

            Framework.ConfigData cd = User.Configuration["日常生活表現總表_假別設定"];
            if (cd.Count != 0)
            {
                XmlElement config = cd.GetXml("XmlData", null);
                if (config.SelectSingleNode("Print") != null)
                    _sizeIndex = int.Parse((config.SelectSingleNode("Print") as XmlElement).GetAttribute("PaperSize").ToString());
            }

            BGW.DoWork += new DoWorkEventHandler(BgW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgW_RunWorkerCompleted);

            BGW_new.DoWork += new DoWorkEventHandler(BGW_new_DoWork);
            BGW_new.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_new_RunWorkerCompleted);
        }

        //int t; //時間偵測

        private void btnReport_Click(object sender, EventArgs e)
        {
            Dictionary<string, List<string>> dic = GetAbsenceItems();
            if (dic.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("假別設定目前為0,需勾選欲列印的假別");
                return;
            }

            if (!BGW.IsBusy && !BGW_new.IsBusy)
            {
                this.Enabled = false;
                SchoolYear = cbSchoolYear.Text;
                Semester = cbSemester.Text;
                IsPredic = checkBoxX1.Checked;
                this.Text = "列印報表中!!";

                if (!checkBoxX1.Checked) //不使用不及格判斷
                {

                    BGW.RunWorkerAsync();
                }
                else //進行不及格判斷
                {
                    BGW_new.RunWorkerAsync();
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("忙碌中請稍後!!");
                this.Enabled = true;
                return;
            }
        }

        //取得學生資料
        void BGW_new_DoWork(object sender, DoWorkEventArgs e)
        {
            List<StudentRecord> StudentList = new List<StudentRecord>();

            List<ClassRecord> ClassList = new List<ClassRecord>();
            ClassList.Clear();
            ClassList = Class.Instance.SelectedList; //班級清單
            ClassList.Sort(new Comparison<ClassRecord>(ClassComparer));

            foreach (ClassRecord each in ClassList)
            {
                foreach (StudentRecord stud in each.Students)
                {
                    if (stud.Status == "一般")
                    {
                        //建立取得資料的學生ID
                        if (!StudentList.Contains(stud))
                        {
                            StudentList.Add(stud);
                        }
                    }
                }
            }
            e.Result = StudentList; //學生ID清單
        }

        private int ClassComparer(ClassRecord x, ClassRecord y)
        {
            string xx = x.Name;
            string yy = y.Name;
            return xx.CompareTo(yy);
        }

        //取得學生資料完成
        void BGW_new_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            List<StudentRecord> StudentList = e.Result as List<StudentRecord>;
            GraduationPredictReportForm GPRF = new GraduationPredictReportForm(StudentList, SchoolYear, Semester);
            DialogResult DR = GPRF.ShowDialog();
            StudentNotPassList = new List<string>();
            if (DR == DialogResult.OK)
            {
                foreach (string each in GPRF._passList.Keys)
                {
                    if (!GPRF._passList[each])
                    {
                        JHSchool.Data.JHStudentRecord sr = JHSchool.Data.JHStudent.SelectByID(each);
                        string GradeYear = sr.Class.GradeYear.HasValue ? sr.Class.GradeYear.Value.ToString() : "0";

                        foreach (ResultDetail detail in GPRF._result[each])
                        {
                            if (detail.Semester == Semester && detail.GradeYear == GradeYear)
                            {
                                if (!StudentNotPassList.Contains(each))
                                {
                                    StudentNotPassList.Add(each);
                                }
                            }
                        }
                    }
                }
                er = GPRF._result;
                BGW.RunWorkerAsync();
            }
            else
            {
                this.Text = "日常生活表現總表(高雄)";
                FISCA.Presentation.Controls.MsgBox.Show("已取消動作!!");
                this.Enabled = true;
            }
        }

        void BgW_DoWork(object sender, DoWorkEventArgs e)
        {
            #region 開始列印

            GetData = new GetMSRlist();
            jjMessage = GetData.GetMSRList(SchoolYear, Semester);

            if (!string.IsNullOrEmpty(jjMessage))
            {
                FISCA.Presentation.Controls.MsgBox.Show("列印提示：\n" + jjMessage);
            }

            //處理資料,整理出依班級為單位之清單
            SuperList = GetData.SuperList;

            //取得學校名稱
            string SchoolName = School.ChineseName;

            ColumnInTitleIndex.Clear();

            ColumnInTitleIndex.Add("座號", 0);
            ColumnInTitleIndex.Add("姓名", 1);
            ColumnInTitleIndex.Add("學號", 2);
            ColumnInTitleIndex.Add("大功", 3);
            ColumnInTitleIndex.Add("小功", 4);
            ColumnInTitleIndex.Add("嘉獎", 5);
            ColumnInTitleIndex.Add("大過", 6);
            ColumnInTitleIndex.Add("小過", 7);
            ColumnInTitleIndex.Add("警告", 8);

            Dictionary<string, List<string>> dic = GetAbsenceItems();

            if (_sizeIndex == 0)
            {
                wb.Open(new MemoryStream(Properties.Resources.日常生活表現總表範本));
            }
            else
            {
                wb.Open(new MemoryStream(Properties.Resources.日常生活表現總表範本B4));
            }

            string SheetTitle = wb.Worksheets["Sheet1"].Name = "日常生活表現總表";

            #region 建立缺曠格式1
            byte AttendanceColumn = 9;
            int temp = AttendanceColumn;

            int a = 0;
            int b = 0;
            int c = 0;
            List<int> Bcount = new List<int>();
            foreach (string each in dic.Keys)
            {
                int countColumn = 0;
                foreach (string eachIn in dic[each])
                {
                    wb.Worksheets[SheetTitle].Cells.InsertColumn(temp);
                    wb.Worksheets["範本"].Cells.InsertColumn(temp);
                    countColumn++;
                    a++; //總共多少
                }
                Bcount.Add(countColumn);
                c++; //假別類別
            }

            int MergeIndex = AttendanceColumn;
            wb.Worksheets[SheetTitle].Cells.Merge(1, MergeIndex, 1, a); //標題
            wb.Worksheets["範本"].Cells.Merge(1, MergeIndex, 1, a); //標題
            for (int x = 0; x < c; x++)
            {
                wb.Worksheets[SheetTitle].Cells.Merge(2, MergeIndex, 1, Bcount[x]); //假別1
                wb.Worksheets["範本"].Cells.Merge(2, MergeIndex, 1, Bcount[x]); //假別1
                MergeIndex += Bcount[x];
            }

            wb.Worksheets[SheetTitle].Cells.DeleteColumn(MergeIndex);
            wb.Worksheets["範本"].Cells.DeleteColumn(MergeIndex);
            #endregion

            #region 填入缺曠標題
            wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].PutValue("缺曠統計資料");
            wb.Worksheets["範本"].Cells[1, AttendanceColumn].PutValue("缺曠統計資料");

            foreach (string each in dic.Keys)
            {
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(each);
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;

                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(each);
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;

                foreach (string eachIn in dic[each])
                {
                    ColumnInTitleIndex.Add(each + eachIn, AttendanceColumn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue(eachIn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].Style.Font.Size = 8; //字型
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].Style.IsTextWrapped = true; //自動換行
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center; //水平置中
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center; //垂直置中
                    wb.Worksheets[SheetTitle].Cells.SetColumnWidth(AttendanceColumn, 2); //column寬度
                    wb.Worksheets[SheetTitle].Cells.SetRowHeight(3, 21.75); //Row高度

                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue(eachIn);
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].Style.Font.Size = 8; //字型
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].Style.IsTextWrapped = true; //自動換行
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center; //水平置中
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center; //垂直置中
                    wb.Worksheets["範本"].Cells.SetColumnWidth(AttendanceColumn, 2); //column寬度
                    wb.Worksheets["範本"].Cells.SetRowHeight(3, 21.75); //Row高度

                    AttendanceColumn++;
                }
            }
            wb.Worksheets[SheetTitle].Cells.CreateRange(3, temp, 1, a).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            wb.Worksheets["範本"].Cells.CreateRange(3, temp, 1, a).SetOutlineBorder(BorderType.TopBorder, CellBorderType.Thin, Color.Black);
            #endregion

            #region 取得日常生活表現設定

            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];

            if (!string.IsNullOrEmpty(cd["DailyBehavior"]))
            {
                XmlElement dailyBehavior = XmlHelper.LoadXml(cd["DailyBehavior"]); //日常生活表現
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(dailyBehavior.GetAttribute("Name"));
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(dailyBehavior.GetAttribute("Name"));

                foreach (XmlElement item in dailyBehavior.SelectNodes("Item"))
                {
                    ColumnInTitleIndex.Add(item.GetAttribute("Name"), AttendanceColumn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    AttendanceColumn++;
                }
            }

            if (!string.IsNullOrEmpty(cd["GroupActivity"]))
            {
                XmlElement groupActivity = XmlHelper.LoadXml(cd["GroupActivity"]); //團體活動
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(groupActivity.GetAttribute("Name"));
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(groupActivity.GetAttribute("Name"));

                //因為社團活動,所以把社團活動手動增加
                //ColumnInTitleIndex.Add("社團活動", AttendanceColumn);
                //wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue("社團活動");
                //wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue("社團活動");
                //AttendanceColumn++;

                foreach (XmlElement item in groupActivity.SelectNodes("Item"))
                {
                    if (item.GetAttribute("Name") == "社團活動" || item.GetAttribute("Name") == "自治活動")
                        continue;

                    ColumnInTitleIndex.Add(item.GetAttribute("Name"), AttendanceColumn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    AttendanceColumn++;
                }
            }

            if (!string.IsNullOrEmpty(cd["PublicService"]))
            {
                XmlElement publicService = XmlHelper.LoadXml(cd["PublicService"]); //公共服務
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(publicService.GetAttribute("Name"));
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(publicService.GetAttribute("Name"));

                foreach (XmlElement item in publicService.SelectNodes("Item"))
                {
                    ColumnInTitleIndex.Add(item.GetAttribute("Name"), AttendanceColumn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    AttendanceColumn++;
                }
            }
            if (!string.IsNullOrEmpty(cd["SchoolSpecial"]))
            {

                XmlElement schoolSpecial = XmlHelper.LoadXml(cd["SchoolSpecial"]); //校內外
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(schoolSpecial.GetAttribute("Name"));
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(schoolSpecial.GetAttribute("Name"));

                foreach (XmlElement item in schoolSpecial.SelectNodes("Item"))
                {
                    ColumnInTitleIndex.Add(item.GetAttribute("Name"), AttendanceColumn);
                    wb.Worksheets[SheetTitle].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    wb.Worksheets["範本"].Cells[3, AttendanceColumn].PutValue(item.GetAttribute("Name"));
                    AttendanceColumn++;
                }
            }

            if (!string.IsNullOrEmpty(cd["DailyLifeRecommend"]))
            {

                XmlElement dailyLifeRecommend = XmlHelper.LoadXml(cd["DailyLifeRecommend"]); //日常生活表現具體建議

                ColumnInTitleIndex.Add(dailyLifeRecommend.Name, AttendanceColumn);
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue(dailyLifeRecommend.GetAttribute("Name"));
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue(dailyLifeRecommend.GetAttribute("Name"));

                wb.Worksheets[SheetTitle].Cells.Merge(2, AttendanceColumn, 2, 1);
                wb.Worksheets["範本"].Cells.Merge(2, AttendanceColumn, 2, 1);
            }

            if (IsPredic) //是否進行及格判斷
            {
                #region 及格判斷
                Range rg999 = wb.Worksheets[SheetTitle].Cells.CreateRange(AttendanceColumn, 1, true);

                AttendanceColumn++;
                ColumnInTitleIndex.Add("不及格判斷", AttendanceColumn);

                wb.Worksheets[SheetTitle].Cells.CreateRange(AttendanceColumn, 1, true).Copy(rg999);
                wb.Worksheets["範本"].Cells.CreateRange(AttendanceColumn, 1, true).Copy(rg999);
                //清空字樣
                wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue("");
                wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue("");
                //合併欄位
                wb.Worksheets[SheetTitle].Cells.Merge(1, AttendanceColumn, 2, 1);
                wb.Worksheets["範本"].Cells.Merge(1, AttendanceColumn, 2, 1);
                //調整欄位大小
                wb.Worksheets[SheetTitle].Cells.SetColumnWidth(AttendanceColumn, 8);
                wb.Worksheets["範本"].Cells.SetColumnWidth(AttendanceColumn, 8);
                //垂直置中
                wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                wb.Worksheets["範本"].Cells[1, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                //水平置中
                wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;
                wb.Worksheets["範本"].Cells[1, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;
                //填入抬頭
                wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].PutValue("不及格判斷");
                wb.Worksheets["範本"].Cells[1, AttendanceColumn].PutValue("不及格判斷");
                #endregion

                #region 及格判斷

                //AttendanceColumn++;
                //ColumnInTitleIndex.Add("不及格說明", AttendanceColumn);

                //wb.Worksheets[SheetTitle].Cells.CreateRange(AttendanceColumn, 1, true).Copy(rg999);
                //wb.Worksheets["範本"].Cells.CreateRange(AttendanceColumn, 1, true).Copy(rg999);

                //清空字樣
                //wb.Worksheets[SheetTitle].Cells[2, AttendanceColumn].PutValue("");
                //wb.Worksheets["範本"].Cells[2, AttendanceColumn].PutValue("");

                //合併欄位
                //wb.Worksheets[SheetTitle].Cells.Merge(1, AttendanceColumn, 2, 1);
                //wb.Worksheets["範本"].Cells.Merge(1, AttendanceColumn, 2, 1);

                //調整欄位大小
                //wb.Worksheets[SheetTitle].Cells.SetColumnWidth(AttendanceColumn, 8);
                //wb.Worksheets["範本"].Cells.SetColumnWidth(AttendanceColumn, 8);

                //垂直置中
                //wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                //wb.Worksheets["範本"].Cells[1, AttendanceColumn].Style.HorizontalAlignment = TextAlignmentType.Center;
                //水平置中
                //wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;
                //wb.Worksheets["範本"].Cells[1, AttendanceColumn].Style.VerticalAlignment = TextAlignmentType.Center;
                //填入抬頭
                //wb.Worksheets[SheetTitle].Cells[1, AttendanceColumn].PutValue("不及格說明");
                //wb.Worksheets["範本"].Cells[1, AttendanceColumn].PutValue("不及格說明");

                #endregion
            }

            #endregion

            //標題合併
            wb.Worksheets[SheetTitle].Cells.Merge(0, 0, 1, AttendanceColumn++);
            wb.Worksheets["範本"].Cells.Merge(0, 0, 1, AttendanceColumn++);
            
            Range RangeDetil = wb.Worksheets["範本"].Cells.CreateRange(0, 4, false);
            Range RangeRow = wb.Worksheets["範本"].Cells.CreateRange(4, 1, false);

            #region 學生資料
            int rowIndex = 0; //RowIndex

            foreach (string each in SuperList.Keys) //取得班級ID
            {
                if (SuperList[each].Count == 0) //如果沒有學生就離開
                    continue;

                wb.Worksheets[SheetTitle].Cells.CreateRange(rowIndex, 4, false).Copy(RangeDetil);

                JHClassRecord CR = JHClass.SelectByID(each);

                StringBuilder sb = new StringBuilder();
                sb.Append(SchoolName + "　");
                sb.Append(SchoolYear + "學年度　");
                sb.Append("第" + Semester + "學期　");
                sb.Append(CR.Name + "班　");
                sb.Append("日常生活表現總表");

                wb.Worksheets[SheetTitle].Cells[rowIndex, 0].PutValue(sb.ToString());
                wb.Worksheets["範本"].Cells[rowIndex, 0].PutValue(sb.ToString());

                //wb.Worksheets[SheetTitle].Cells[rowIndex + 1, 2].PutValue(CR.Name); //填入班級名稱

                rowIndex += 4;

                foreach (string stud in SuperList[each].Keys) //取得學生ID
                {
                    string StudAssn = GetAssnString(stud);

                    wb.Worksheets[SheetTitle].Cells.CreateRange(rowIndex, 1, false).Copy(RangeRow);

                    JHStudentRecord SR = JHStudent.SelectByID(stud); //取得學生

                    JHMoralScoreRecord MSR = SuperList[each][stud]; //取得日常生活表現

                    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["座號"]].PutValue(SR.SeatNo);
                    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["姓名"]].PutValue(SR.Name);
                    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["學號"]].PutValue(SR.StudentNumber);

                    if (MSR != null)
                    {
                        #region 日常生活表現

                        Dictionary<string, string> BehaviorDic = new Dictionary<string, string>();

                        if (MSR.TextScore != null)
                        {
                            XmlElement DailyBehavioNode = (XmlElement)MSR.TextScore.SelectSingleNode("DailyBehavior");

                            if (DailyBehavioNode != null)
                            {
                                foreach (XmlElement item in DailyBehavioNode.SelectNodes("Item"))
                                {
                                    if (ColumnInTitleIndex.ContainsKey(item.GetAttribute("Name")))
                                    {
                                        BehaviorDic.Add(item.GetAttribute("Name"), item.GetAttribute("Degree"));
                                    }
                                }
                            }

                            XmlElement GroupActivityNode = (XmlElement)MSR.TextScore.SelectSingleNode("GroupActivity");

                            //BehaviorDic.Add("社團活動", StudAssn);

                            if (GroupActivityNode != null)
                            {
                                foreach (XmlElement item in GroupActivityNode.SelectNodes("Item"))
                                {
                                    if (item.GetAttribute("Name") == "社團活動" || item.GetAttribute("Name") == "自治活動")
                                        continue;

                                    if (ColumnInTitleIndex.ContainsKey(item.GetAttribute("Name")))
                                    {
                                        if (item.GetAttribute("Degree") + item.GetAttribute("Description") != "")
                                        {
                                            BehaviorDic.Add(item.GetAttribute("Name"), item.GetAttribute("Degree") + "," + item.GetAttribute("Description"));
                                        }
                                        else
                                        {
                                            BehaviorDic.Add(item.GetAttribute("Name"), "");
                                        }
                                    }
                                }
                            }

                            XmlElement PublicServiceNode = (XmlElement)MSR.TextScore.SelectSingleNode("PublicService");

                            if (PublicServiceNode != null)
                            {
                                foreach (XmlElement item in PublicServiceNode.SelectNodes("Item"))
                                {
                                    if (ColumnInTitleIndex.ContainsKey(item.GetAttribute("Name")))
                                    {
                                        BehaviorDic.Add(item.GetAttribute("Name"), item.GetAttribute("Description"));
                                    }
                                }
                            }

                            XmlElement SchoolSpecialNode = (XmlElement)MSR.TextScore.SelectSingleNode("SchoolSpecial");

                            if (SchoolSpecialNode != null)
                            {
                                foreach (XmlElement item in SchoolSpecialNode.SelectNodes("Item"))
                                {
                                    if (ColumnInTitleIndex.ContainsKey(item.GetAttribute("Name")))
                                    {
                                        BehaviorDic.Add(item.GetAttribute("Name"), item.GetAttribute("Description"));
                                    }
                                }
                            }

                            XmlElement DailyLifeRecommendNode = (XmlElement)MSR.TextScore.SelectSingleNode("DailyLifeRecommend");

                            if (DailyLifeRecommendNode != null)
                            {
                                if (ColumnInTitleIndex.ContainsKey(DailyLifeRecommendNode.Name))
                                {
                                    BehaviorDic.Add(DailyLifeRecommendNode.Name, DailyLifeRecommendNode.GetAttribute("Description"));
                                }
                            }
                        }


                        #endregion

                        foreach (string Beh in BehaviorDic.Keys)
                        {
                            wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex[Beh]].PutValue(BehaviorDic[Beh]);
                        }
                    }
                    //else
                    //{
                    //    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["座號"]].PutValue(SR.SeatNo);
                    //    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["姓名"]].PutValue(SR.Name);
                    //    wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["學號"]].PutValue(SR.StudentNumber);
                    //    rowIndex++;
                    //}

                    //取得自動統計內容
                    AutoSummaryRecord ASRObj = GetData.AutoSummaryDic[stud];

                    if (ASRObj != null)
                    {
                        #region 缺曠獎懲統計資料

                        string MeritA = "";
                        string MeritB = "";
                        string MeritC = "";
                        string DemeritA = "";
                        string DemeritB = "";
                        string DemeritC = "";

                        Dictionary<string, string> AttendanceDic = new Dictionary<string, string>();

                        AutoSummaryRecord AutoSummaryObj = ASRObj;

                        #region 獎懲資料
                        XmlElement MeritItem = (XmlElement)AutoSummaryObj.AutoSummary.SelectSingleNode("DisciplineStatistics/Merit");
                        if (MeritItem != null)
                        {
                            MeritA = MeritItem.GetAttribute("A");
                            MeritB = MeritItem.GetAttribute("B");
                            MeritC = MeritItem.GetAttribute("C");
                        }

                        XmlElement DemeritItem = (XmlElement)AutoSummaryObj.AutoSummary.SelectSingleNode("DisciplineStatistics/Demerit");

                        if (DemeritItem != null)
                        {
                            DemeritA = DemeritItem.GetAttribute("A");
                            DemeritB = DemeritItem.GetAttribute("B");
                            DemeritC = DemeritItem.GetAttribute("C");
                        }
                        #endregion

                        #region 缺曠資料

                        XmlNodeList nodeList = AutoSummaryObj.AutoSummary.SelectNodes("AttendanceStatistics/Absence");

                        foreach (XmlNode itemNode in nodeList)
                        {
                            XmlElement itemElement = (XmlElement)itemNode;
                            string name = itemElement.GetAttribute("Name");
                            string periodtype = itemElement.GetAttribute("PeriodType");
                            int count;
                            if (!int.TryParse(itemElement.GetAttribute("Count"), out count))
                                count = 0;

                            if (ColumnInTitleIndex.ContainsKey(periodtype + name))
                            {
                                AttendanceDic.Add(periodtype + name, count == 0 ? "" : count.ToString());
                            }

                        }

                        #endregion

                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["大功"]].PutValue(MeritA == "0" ? "" : MeritA);
                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["小功"]].PutValue(MeritB == "0" ? "" : MeritB);
                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["嘉獎"]].PutValue(MeritC == "0" ? "" : MeritC);
                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["大過"]].PutValue(DemeritA == "0" ? "" : DemeritA);
                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["小過"]].PutValue(DemeritB == "0" ? "" : DemeritB);
                        wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["警告"]].PutValue(DemeritC == "0" ? "" : DemeritC);

                        foreach (string att in AttendanceDic.Keys)
                        {
                            wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex[att]].PutValue(AttendanceDic[att]);
                        }

                        #endregion
                    }

                    if (IsPredic)
                    {
                        if (!StudentNotPassList.Contains(stud)) //不及格學生
                        {

                        }
                        else //不及格
                        {
                            wb.Worksheets[SheetTitle].Cells[rowIndex, ColumnInTitleIndex["不及格判斷"]].PutValue("不及格");
                        }
                    }

                    rowIndex++;
                }
                wb.Worksheets[SheetTitle].HPageBreaks.Add(rowIndex, ColumnInTitleIndex.Count);
            }
            #endregion
            wb.Worksheets[SheetTitle].AutoFitRows();
            wb.Worksheets.RemoveAt("範本");

            #endregion
        }

        /// <summary>
        /// 傳入學生ID,取得社團字串
        /// </summary>
        /// <param name="stud"></param>
        /// <returns></returns>
        private string GetAssnString(string stud)
        {
            #region 傳入學生ID,取得社團字串
            string AssnScore = "";

            if (!GetData.AssnList.ContainsKey(stud))
            {
                return AssnScore;
            }
            AssnCode assn = GetData.AssnList[stud];


            if (assn.Scores == "")
                return AssnScore;

            XmlElement xml = DSXmlHelper.LoadXml(assn.Scores);


            foreach (XmlElement each in xml.SelectNodes("Item"))
            {
                //string x0 = each.GetAttribute("AssociationName");
                string x1 = each.GetAttribute("Score");
                string x2 = each.GetAttribute("Effort");
                string x3 = each.GetAttribute("Text");

                //if (x0 != "")
                //{
                //    AssnScore += "" + x0 + "。";
                //}

                if (x1 != "")
                {
                    AssnScore += "" + x1 + "。";
                }

                if (x2 != "")
                {
                    int jj;
                    if (int.TryParse(x2, out jj))
                    {
                        AssnScore += "" + Effor.GetTextByCode(jj) + "。";
                    }
                    else
                    {
                        AssnScore += "" + x2 + "。";
                    }
                }

                if (x3 != "")
                {
                    AssnScore += "" + x3 + "。";
                }
            }

            return AssnScore;
            #endregion
        }

        void BgW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            #region 背景工作完成後...

            this.Text = "日常生活表現總表(高雄)";
            this.Enabled = true;

            if (e.Error == null)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("日常生活表現總表,列印完成!!");

                SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = "日常生活表現總表.xls";
                sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.Save(sd.FileName, FileFormatType.Excel2003);
                        System.Diagnostics.Process.Start(sd.FileName);

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        this.Enabled = true;
                        return;
                    }
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("列印發生錯誤!!\n" + e.Error.Message);
            }


            #endregion
        }

        /// <summary>
        /// 取得假別項目
        /// </summary>
        /// <returns>假別項目列表</returns>
        public static Dictionary<string, List<string>> GetAbsenceItems()
        {
            #region 取得假別項目
            //讀取缺曠別 Preference
            Dictionary<string, List<string>> config = new Dictionary<string, List<string>>();

            //XmlElement preferenceData = CurrentUser.Instance.Preference["缺曠通知單_缺曠別設定"];
            Framework.ConfigData cd = User.Configuration["日常生活表現總表_假別設定"];
            XmlElement preferenceData = cd.GetXml("XmlData", null);

            if (preferenceData != null)
            {
                foreach (XmlElement type in preferenceData.SelectNodes("Type"))
                {
                    string prefix = type.GetAttribute("Text");
                    if (!config.ContainsKey(prefix))
                        config.Add(prefix, new List<string>());

                    foreach (XmlElement absence in type.SelectNodes("Absence"))
                    {
                        if (!config[prefix].Contains(absence.GetAttribute("Text")))
                            config[prefix].Add(absence.GetAttribute("Text"));
                    }
                }
            }

            return config;
            #endregion
        }

        private int SrotSeanNum(JHStudentRecord xx, JHStudentRecord yy)
        {
            #region 依座號排序
            string x;
            string y;
            if (xx.SeatNo.HasValue)
            {
                x = xx.SeatNo.ToString();
            }
            else
            {
                x = "";
            }

            if (yy.SeatNo.HasValue)
            {
                y = yy.SeatNo.ToString();
            }
            else
            {
                y = "";
            }

            return x.CompareTo(y);
            #endregion
        }

        #region 其他處理程序
        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SelectTypeForm STF = new SelectTypeForm("日常生活表現總表_假別設定");
            STF.ShowDialog();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            WeekAbsenceReportConfig WKRC = new WeekAbsenceReportConfig("日常生活表現總表_假別設定", _sizeIndex);
            WKRC.ShowDialog();
            _sizeIndex = WKRC.SizeIndex;
        }
        #endregion

        private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                btnReport.Text = "下一步";
            }
            else
            {
                btnReport.Text = "列印";
            }
        }
    }
}
