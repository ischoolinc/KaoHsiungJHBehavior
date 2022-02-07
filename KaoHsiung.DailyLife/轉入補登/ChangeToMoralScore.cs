using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using JHSchool.Data;
using System.Xml;
using Framework;
using FISCA.DSAUtil;
using JHSchool;
using Framework.Feature;
using K12.Data.Utility;
using FISCA.LogAgent;
using JHSchool.Behavior.BusinessLogic;
using JHSchool.Behavior;
using Campus.Windows;

namespace KaoHsiung.DailyLife
{
    public partial class ChangeToMoralScore : BaseForm
    {
        private JHMoralScoreRecord _editorRecord; //修改模式使用
        private string _RefStudentID; //新增模式使用

        private Dictionary<string, string> Morality = new Dictionary<string, string>(); //日常生活表現具體建議使用
        private Dictionary<string, string> EffortList = new Dictionary<string, string>();  //努力程度代碼
        private Dictionary<string, string> dic = new Dictionary<string, string>();

        //用來記錄 日常行為表現 是否有格子輸入不正確的值
        private Dictionary<DataGridViewCell, bool> inputErrors = new Dictionary<DataGridViewCell, bool>();

        private List<string> _periodTypes;
        private List<string> _absenceList;
        private List<string> _meritTypes;

        private Dictionary<string, int> periodList;
        private Dictionary<string, int> meritList;
        private string Mode;

        /// <summary>
        /// 轉入補登新增模式
        /// </summary>
        /// <param name="RefID"></param>
        public ChangeToMoralScore(string RefID)
        {
            InitializeComponent();

            _RefStudentID = RefID;

            Mode = "NEW";

            this.Text = "轉入補登(新增)";

            dgvDailyLifeRecommend.Rows.Add();

            #region 學年度學期
            string schoolYear = School.DefaultSchoolYear;
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 4).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 3).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 1).ToString());
            int x = cbSchoolYear.Items.Add((int.Parse(schoolYear)).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 1).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 3).ToString());
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

            SyndLoad(); //建立預設畫面
        }

        /// <summary>
        /// 轉入補登修改模式
        /// </summary>
        /// <param name="MSR"></param>
        public ChangeToMoralScore(JHMoralScoreRecord MSR)
        {
            InitializeComponent();

            Mode = "UPDATA";

            this.Text = "轉入補登(修改)";

            cbSchoolYear.Items.Add(MSR.SchoolYear);
            cbSchoolYear.SelectedIndex = 0;
            cbSemester.Items.Add(MSR.Semester);
            cbSemester.SelectedIndex = 0;

            dgvDailyLifeRecommend.Rows.Add();

            _editorRecord = MSR;

            cbSchoolYear.Text = _editorRecord.SchoolYear.ToString();
            cbSemester.Text = _editorRecord.Semester.ToString();
            cbSchoolYear.Enabled = false;
            cbSemester.Enabled = false;

            SyndLoad(); //建立預設畫面

            BindData(); //填入資料
        }

        /// <summary>
        /// 建立畫面資料
        /// </summary>
        /// <param name="MSRData"></param>
        private void SyndLoad()
        {
            #region 建立畫面資料1
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];

            if (cd.Contains("DailyBehavior"))
            {
                XmlElement dailyBehavior = XmlHelper.LoadXml(cd["DailyBehavior"]);
                tabControl1.Tabs[0].Text = dailyBehavior.GetAttribute("Name");

                foreach (XmlElement item in dailyBehavior.SelectNodes("Item"))
                    dgvDailyBehavior.Rows.Add(item.GetAttribute("Name"), item.GetAttribute("Index"), "");
            }

            if (cd.Contains("GroupActivity"))
            {
                XmlElement groupActivity = XmlHelper.LoadXml(cd["GroupActivity"]);
                tabControl1.Tabs[1].Text = groupActivity.GetAttribute("Name");

                foreach (XmlElement item in groupActivity.SelectNodes("Item"))
                    dgvGroupActivity.Rows.Add(item.GetAttribute("Name"), "", "");
            }

            if (cd.Contains("PublicService"))
            {
                XmlElement publicService = XmlHelper.LoadXml(cd["PublicService"]);
                tabControl1.Tabs[2].Text = publicService.GetAttribute("Name");

                foreach (XmlElement item in publicService.SelectNodes("Item"))
                    dgvPublicService.Rows.Add(item.GetAttribute("Name"), "");
            }

            if (cd.Contains("SchoolSpecial"))
            {
                XmlElement schoolSpecial = XmlHelper.LoadXml(cd["SchoolSpecial"]);
                tabControl1.Tabs[3].Text = schoolSpecial.GetAttribute("Name");

                foreach (XmlElement item in schoolSpecial.SelectNodes("Item"))
                    dgvSchoolSpecial.Rows.Add(item.GetAttribute("Name"), "");
            }

            if (cd.Contains("DailyLifeRecommend"))
            {
                XmlElement dailyLifeRecommend = XmlHelper.LoadXml(cd["DailyLifeRecommend"]);
                tabControl1.Tabs[4].Text = dailyLifeRecommend.GetAttribute("Name");
            }

            //努力程度
            ReflashEffortList();
            //日常行為表現
            ReflashDic();

            //日常生活表現具體建議使用
            ReflashMorality();
            #endregion

            #region 建立畫面資料2
            dataGridViewX3.Rows.Clear();
            dataGridViewX4.Rows.Clear();

            _periodTypes = GetPeriodTypeItems();
            _periodTypes.Sort();
            _absenceList = GetAbsenceItems();
            _meritTypes = GetMeritTypes();

            #region 將節次類別,填入欄位

            periodList = new Dictionary<string, int>();

            foreach (string periodType in _periodTypes)
            {
                foreach (string each in _absenceList)
                {
                    string columnName = periodType + each;
                    int periodIndex = dataGridViewX3.Columns.Add(columnName, columnName);

                    dataGridViewX3.Columns[columnName].Width = 80;
                    dataGridViewX3.Columns[columnName].Tag = periodType + ":" + each;
                    dataGridViewX3.Columns[columnName].SortMode = DataGridViewColumnSortMode.NotSortable;

                    periodList.Add(columnName, periodIndex);
                }
            }
            #endregion

            #region 塞入獎懲之預設值
            meritList = new Dictionary<string, int>();
            foreach (string merit in _meritTypes)
            {
                int columnIndex = dataGridViewX4.Columns.Add(merit, merit);

                dataGridViewX4.Columns[columnIndex].Width = 110;
                dataGridViewX4.Columns[columnIndex].SortMode = DataGridViewColumnSortMode.NotSortable;

                meritList.Add(merit, columnIndex);
            }
            #endregion

            dataGridViewX3.Rows.Add();
            dataGridViewX4.Rows.Add();
            #endregion

            #region 處理dataGridView可填入數字部分轉半形英數

            DataGridViewImeDecorator decX3 = new DataGridViewImeDecorator(this.dataGridViewX3);

            DataGridViewImeDecorator decX4 = new DataGridViewImeDecorator(this.dataGridViewX4);

            List<int> colsDgvGroupActivity = new List<int>() { 1 };
            DataGridViewImeDecorator dec = new DataGridViewImeDecorator(this.dgvGroupActivity, colsDgvGroupActivity);

            List<int> colsDgvDailyBehavior = new List<int>() { 2 };
            DataGridViewImeDecorator dec2 = new DataGridViewImeDecorator(this.dgvDailyBehavior, colsDgvDailyBehavior);
            #endregion
        }

        private void BindData()
        {
            #region 更新資料1
            XmlElement node;

            if (_editorRecord == null) return;

            //DailyBehavior
            //<DailyBehavior Name="日常行為表現">
            //    <Item Name="愛整潔" Index="....." Degree="3"/>
            //    <Item Name="守秩序" Index="....." Degree="3"/>
            //</DailyBehavior>
            node = (XmlElement)_editorRecord.TextScore.SelectSingleNode("DailyBehavior");

            if (node != null)
            {
                foreach (XmlElement item in node.SelectNodes("Item"))
                {
                    foreach (DataGridViewRow row in dgvDailyBehavior.Rows)
                    {
                        if (row.Cells[0].Value.ToString() == item.GetAttribute("Name"))
                        {
                            row.Cells[1].Value = item.GetAttribute("Index");
                            row.Cells[2].Value = item.GetAttribute("Degree");
                        }
                    }
                }
            }

            //GroupActivity
            //<GroupActivity Name="團體活動表現">
            //    <Item Name="社團活動" Degree="1" Description=".....">
            //    <Item Name="學校活動" Degree="2" Description=".....">
            //</GroupActivity>
            node = (XmlElement)_editorRecord.TextScore.SelectSingleNode("GroupActivity");

            if (node != null)
            {
                foreach (XmlElement item in node.SelectNodes("Item"))
                {
                    foreach (DataGridViewRow row in dgvGroupActivity.Rows)
                    {
                        if (row.Cells[0].Value.ToString() == item.GetAttribute("Name"))
                        {
                            row.Cells[1].Value = item.GetAttribute("Degree");
                            row.Cells[2].Value = item.GetAttribute("Description");
                        }
                    }
                }
            }

            //PublicService
            //<PublicService Name="公共服務表現">
            //    <Item Name="校內服務" Description=".....">
            //    <Item Name="社區服務" Description=".....">
            //</PublicService>
            node = (XmlElement)_editorRecord.TextScore.SelectSingleNode("PublicService");

            if (node != null)
            {
                foreach (XmlElement item in node.SelectNodes("Item"))
                {
                    foreach (DataGridViewRow row in dgvPublicService.Rows)
                    {
                        if (row.Cells[0].Value.ToString() == item.GetAttribute("Name"))
                        {
                            row.Cells[1].Value = item.GetAttribute("Description");
                        }
                    }
                }
            }

            //SchoolSpecial
            //<SchoolSpecial Name="校內外時特殊表現">
            //    <Item Name="校外特殊表現" Description=".....">
            //    <Item Name="校內特殊表現" Description=".....">
            //</SchoolSpecial>
            node = (XmlElement)_editorRecord.TextScore.SelectSingleNode("SchoolSpecial");

            if (node != null)
            {
                foreach (XmlElement item in node.SelectNodes("Item"))
                {
                    foreach (DataGridViewRow row in dgvSchoolSpecial.Rows)
                    {
                        if (row.Cells[0].Value.ToString() == item.GetAttribute("Name"))
                        {
                            row.Cells[1].Value = item.GetAttribute("Description");
                        }
                    }
                }
            }

            //DailyLifeRecommend
            //<DailyLifeRecommend Name="日常生活表現具體建議" Description=".....">
            node = (XmlElement)_editorRecord.TextScore.SelectSingleNode("DailyLifeRecommend");

            if (node != null)
                dgvDailyLifeRecommend.Rows[0].Cells[0].Value = node.GetAttribute("Description");
            #endregion

            #region 更新資料InitialSummary
            //先清空
            dataGridViewX3.Rows.Clear();
            dataGridViewX4.Rows.Clear();

            //再Add新Row
            dataGridViewX3.Rows.Add();
            dataGridViewX4.Rows.Add();

            SchoolYearSemester SSYear = new SchoolYearSemester(_editorRecord.SchoolYear, _editorRecord.Semester);
            List<AutoSummaryRecord> AsummaryList = AutoSummary.Select(new string[] { _editorRecord.RefStudentID }, new SchoolYearSemester[] { SSYear });

            if (AsummaryList.Count == 1)
            {
                XmlElement NewXml2 = AsummaryList[0].AutoSummary;

                XmlNode now1 = NewXml2.SelectSingleNode("AttendanceStatistics");

                if (now1 != null)
                {
                    foreach (XmlElement each in now1.SelectNodes("Absence"))
                    {
                        string x1 = each.GetAttribute("PeriodType");
                        string x2 = each.GetAttribute("Name");
                        string x3 = each.GetAttribute("Count");

                        if (periodList.ContainsKey(x1 + x2))
                        {
                            dataGridViewX3.Rows[0].Cells[periodList[x1 + x2]].Value = x3;
                        }
                    }
                }

                XmlNode now2 = NewXml2.SelectSingleNode("DisciplineStatistics");
                if (now2 != null)
                {
                    XmlElement BING1 = (XmlElement)now2.SelectSingleNode("Merit");
                    if (BING1 != null)
                    {
                        dataGridViewX4.Rows[0].Cells[meritList["大功"]].Value = BING1.GetAttribute("A");
                        dataGridViewX4.Rows[0].Cells[meritList["小功"]].Value = BING1.GetAttribute("B");
                        dataGridViewX4.Rows[0].Cells[meritList["嘉獎"]].Value = BING1.GetAttribute("C");
                    }

                    XmlElement BING2 = (XmlElement)now2.SelectSingleNode("Demerit");
                    if (BING2 != null)
                    {
                        dataGridViewX4.Rows[0].Cells[meritList["大過"]].Value = BING2.GetAttribute("A");
                        dataGridViewX4.Rows[0].Cells[meritList["小過"]].Value = BING2.GetAttribute("B");
                        dataGridViewX4.Rows[0].Cells[meritList["警告"]].Value = BING2.GetAttribute("C");
                    }
                }
            }
            #endregion
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            if (CheckErrorText())
            {
                if (Mode == "NEW")
                {
                    #region NEW

                    List<JHMoralScoreRecord> MSRList = JHMoralScore.SelectByStudentIDs(new string[] { _RefStudentID });

                    foreach (JHMoralScoreRecord each in MSRList)
                    {
                        if (each.SchoolYear.ToString() == cbSchoolYear.Text && each.Semester.ToString() == cbSemester.Text)
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("該學年度學期資料已經存在,無法新增");
                            return;
                        }
                    }

                    JHMoralScoreRecord NewMsr = newSaveData(); //新增模式

                    try
                    {
                        JHMoralScore.Insert(NewMsr);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("新增資料錯誤");
                        return;
                    }

                    #region Log(新增部份)
                    JHStudentRecord JHSR = JHStudent.SelectByID(_RefStudentID);

                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("詳細資料：");
                    if (JHSR.Class != null)
                    {
                        sb.AppendLine("學生「" + JHSR.Name + "」班級「" + JHSR.Class.Name + "」座號「" + JHSR.SeatNo + "」學號「" + JHSR.StudentNumber + "」。");
                    }
                    else
                    {
                        sb.AppendLine("學生「" + JHSR.Name + "」學號「" + JHSR.StudentNumber + "」。");
                    }

                    sb.AppendLine("學年度「" + cbSchoolYear.Text + "」學期「" + cbSemester.Text + "」");

                    ApplicationLog.Log("日常生活表現模組.轉入補登", "新增日常生活表現資料", "student", JHSR.ID, "由「轉入補登」功能，新增「日常生活表現」資料。\n" + sb.ToString());

                    #endregion

                    _editorRecord = NewMsr;
                    cbSchoolYear.Enabled = false;
                    cbSemester.Enabled = false;
                    Mode = "UPDATA";

                    //新竹模組也檢查???
                    FISCA.Presentation.Controls.MsgBox.Show("新增資料成功");
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                    #endregion

                }
                else if (Mode == "UPDATA")
                {
                    #region UPDATA

                    updataSaveData();

                    try
                    {
                        JHMoralScore.Update(_editorRecord);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("更新資料錯誤");
                        return;
                    }

                    #region Log(更新部份)
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("詳細資料：");
                    if (_editorRecord.Student.Class != null)
                    {
                        sb.AppendLine("學生「" + _editorRecord.Student.Name + "」班級「" + _editorRecord.Student.Class.Name + "」座號「" + _editorRecord.Student.SeatNo + "」學號「" + _editorRecord.Student.StudentNumber + "」。");
                    }
                    else
                    {
                        sb.AppendLine("學生「" + _editorRecord.Student.Name + "」學號「" + _editorRecord.Student.StudentNumber + "」。");
                    }

                    sb.AppendLine("學年度「" + _editorRecord.SchoolYear.ToString() + "」學期「" + _editorRecord.Semester.ToString() + "」");

                    ApplicationLog.Log("日常生活表現模組.轉入補登", "修改日常生活表現資料", "student", _editorRecord.Student.ID, "由「轉入補登」功能，修改「日常生活表現」資料。\n" + sb.ToString());

                    #endregion

                    FISCA.Presentation.Controls.MsgBox.Show("更新資料成功");
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                    #endregion
                }
                else
                {
                    return;
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料有誤,請檢查資料正確性!");
            }
        }

        /// <summary>
        /// 建立新增資料
        /// </summary>
        private JHMoralScoreRecord newSaveData()
        {
            #region 新增模式
            JHMoralScoreRecord msr = new JHMoralScoreRecord();
            msr.RefStudentID = _RefStudentID;
            msr.SchoolYear = int.Parse(cbSchoolYear.Text);
            msr.Semester = int.Parse(cbSemester.Text);

            msr.TextScore = TextScoreData();
            //msr.Summary = SummaryData();
            msr.InitialSummary = InitialSummaryData();
            return msr;

            #endregion
        }

        /// <summary>
        /// 建立更新資料
        /// </summary>
        private void updataSaveData()
        {
            #region 更新模式

            _editorRecord.TextScore = TextScoreData();
            //_editorRecord.Summary = SummaryData();
            _editorRecord.InitialSummary = InitialSummaryData();

            #endregion
        }

        private XmlElement TextScoreData()
        {
            #region 處理TextScore資料
            DSXmlHelper helper = new DSXmlHelper("TextScore");

            //DailyBehavior
            //<DailyBehavior Name="日常行為表現">
            //    <Item Name="愛整潔" Index="....." Degree="3"/>
            //    <Item Name="守秩序" Index="....." Degree="3"/>
            //</DailyBehavior>
            helper.AddElement("DailyBehavior").SetAttribute("Name", tabControl1.Tabs[0].Text);
            foreach (DataGridViewRow row in dgvDailyBehavior.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("DailyBehavior", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Index", "" + row.Cells[1].Value);
                node.SetAttribute("Degree", "" + row.Cells[2].Value);
            }

            //GroupActivity
            //<GroupActivity Name="團體活動表現">
            //    <Item Name="社團活動" Degree="1" Description=".....">
            //    <Item Name="學校活動" Degree="2" Description=".....">
            //</GroupActivity>
            helper.AddElement("GroupActivity").SetAttribute("Name", tabControl1.Tabs[1].Text);
            foreach (DataGridViewRow row in dgvGroupActivity.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("GroupActivity", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Degree", "" + row.Cells[1].Value);
                node.SetAttribute("Description", "" + row.Cells[2].Value);
            }

            //PublicService
            //<PublicService Name="公共服務表現">
            //    <Item Name="校內服務" Description=".....">
            //    <Item Name="社區服務" Description=".....">
            //</PublicService>
            helper.AddElement("PublicService").SetAttribute("Name", tabControl1.Tabs[2].Text);
            foreach (DataGridViewRow row in dgvPublicService.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("PublicService", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Description", "" + row.Cells[1].Value);
            }

            //SchoolSpecial
            //<SchoolSpecial Name="校內外時特殊表現">
            //    <Item Name="校外特殊表現" Description=".....">
            //    <Item Name="校內特殊表現" Description=".....">
            //</SchoolSpecial>
            helper.AddElement("SchoolSpecial").SetAttribute("Name", tabControl1.Tabs[3].Text);
            foreach (DataGridViewRow row in dgvSchoolSpecial.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("SchoolSpecial", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Description", "" + row.Cells[1].Value);
            }

            //DailyLifeRecommend
            //<DailyLifeRecommend Name="日常生活表現具體建議" Description=".....">
            XmlElement anode = helper.AddElement("DailyLifeRecommend");
            anode.SetAttribute("Name", tabControl1.Tabs[4].Text);
            anode.SetAttribute("Description", "" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value);

            return helper.BaseElement;

            #endregion
        }

        /// <summary>
        /// 處理Summary資料(註解)
        /// </summary>
        //private XmlElement SummaryData()
        //{
        //DSXmlHelper newInitialSummary = new DSXmlHelper("Summary");

        //newInitialSummary.AddElement("AttendanceStatistics");

        //#region 缺曠
        //foreach (string each1 in _periodTypes) //一般集會
        //{
        //    foreach (string each2 in _absenceList)
        //    {
        //        //new
        //        string PeriodTypesAndAbsenceList = "" + dataGridViewX1.Rows[0].Cells[periodList[each1 + each2]].Value;
        //        if (PeriodTypesAndAbsenceList == "")
        //            continue;

        //        int IntPeriod;
        //        if (!int.TryParse(PeriodTypesAndAbsenceList, out IntPeriod))
        //            continue;

        //        XmlElement xxx = newInitialSummary.AddElement("AttendanceStatistics", "Absence");
        //        xxx.SetAttribute("PeriodType", each1);
        //        xxx.SetAttribute("Name", each2);

        //        string yyy = PeriodTypesAndAbsenceList;
        //        xxx.SetAttribute("Count", yyy);
        //    }
        //} 
        //#endregion

        //#region 獎勵
        //newInitialSummary.AddElement("DisciplineStatistics");
        //XmlElement zzz = newInitialSummary.AddElement("DisciplineStatistics", "Merit");
        //string a = (string)dataGridViewX2.Rows[0].Cells[meritList["大功"]].Value;
        //string b = (string)dataGridViewX2.Rows[0].Cells[meritList["小功"]].Value;
        //string c = (string)dataGridViewX2.Rows[0].Cells[meritList["嘉獎"]].Value;

        ////new
        //int meritInt;
        //if (a == null) //如果是空值
        //{
        //    zzz.SetAttribute("A", "0");
        //}
        //else if (int.TryParse(a, out meritInt)) //如果是數字
        //{
        //    zzz.SetAttribute("A", a);
        //}

        //if (b == null) //如果是空值
        //{
        //    zzz.SetAttribute("B", "0");
        //}
        //else if (int.TryParse(b, out meritInt)) //如果是數字
        //{
        //    zzz.SetAttribute("B", b);
        //}

        //if (c == null) //如果是空值
        //{
        //    zzz.SetAttribute("C", "0");
        //}
        //else if (int.TryParse(c, out meritInt)) //如果是數字
        //{
        //    zzz.SetAttribute("C", c);
        //} 
        //#endregion

        //#region 懲戒
        //XmlElement kkk = newInitialSummary.AddElement("DisciplineStatistics", "Demerit");
        //string u = (string)dataGridViewX2.Rows[0].Cells[meritList["大過"]].Value;
        //string g = (string)dataGridViewX2.Rows[0].Cells[meritList["小過"]].Value;
        //string p = (string)dataGridViewX2.Rows[0].Cells[meritList["警告"]].Value;

        //int demeritInt;
        //if (u == null) //如果是空值
        //{
        //    kkk.SetAttribute("A", "0");
        //}
        //else if (int.TryParse(u, out demeritInt)) //如果是數字
        //{
        //    kkk.SetAttribute("A", u);
        //}

        //if (g == null) //如果是空值
        //{
        //    kkk.SetAttribute("B", "0");
        //}
        //else if (int.TryParse(g, out demeritInt)) //如果是數字
        //{
        //    kkk.SetAttribute("B", g);
        //}

        //if (p == null) //如果是空值
        //{
        //    kkk.SetAttribute("C", "0");
        //}
        //else if (int.TryParse(p, out demeritInt)) //如果是數字
        //{
        //    kkk.SetAttribute("C", p);
        //} 
        //#endregion

        //return newInitialSummary.BaseElement;
        //}

        /// <summary>
        /// 處理InitialSummary資料
        /// </summary>
        private XmlElement InitialSummaryData()
        {
            #region 處理畫面上的資料
            DSXmlHelper newInitialSummary = new DSXmlHelper("InitialSummary");

            newInitialSummary.AddElement("AttendanceStatistics");

            #region 缺曠
            foreach (string each1 in _periodTypes) //一般集會
            {
                foreach (string each2 in _absenceList)
                {
                    //new
                    string PeriodTypesAndAbsenceList = "" + dataGridViewX3.Rows[0].Cells[periodList[each1 + each2]].Value;
                    if (PeriodTypesAndAbsenceList == "")
                        continue;

                    int IntPeriod;
                    if (!int.TryParse(PeriodTypesAndAbsenceList, out IntPeriod))
                        continue;

                    XmlElement xxx = newInitialSummary.AddElement("AttendanceStatistics", "Absence");
                    xxx.SetAttribute("PeriodType", each1);
                    xxx.SetAttribute("Name", each2);

                    string yyy = PeriodTypesAndAbsenceList;
                    xxx.SetAttribute("Count", yyy);
                }
            }
            #endregion

            newInitialSummary.AddElement("DisciplineStatistics");

            #region 獎勵
            XmlElement zzz = newInitialSummary.AddElement("DisciplineStatistics", "Merit");
            string a = "" + dataGridViewX4.Rows[0].Cells[meritList["大功"]].Value;
            string b = "" + dataGridViewX4.Rows[0].Cells[meritList["小功"]].Value;
            string c = "" + dataGridViewX4.Rows[0].Cells[meritList["嘉獎"]].Value;

            //new
            int meritInt;
            if (a == null) //如果是空值
            {
                zzz.SetAttribute("A", "0");
            }
            else if (int.TryParse(a, out meritInt)) //如果是數字
            {
                zzz.SetAttribute("A", a);
            }

            if (b == null) //如果是空值
            {
                zzz.SetAttribute("B", "0");
            }
            else if (int.TryParse(b, out meritInt)) //如果是數字
            {
                zzz.SetAttribute("B", b);
            }

            if (c == null) //如果是空值
            {
                zzz.SetAttribute("C", "0");
            }
            else if (int.TryParse(c, out meritInt)) //如果是數字
            {
                zzz.SetAttribute("C", c);
            }
            #endregion

            #region 懲戒
            XmlElement kkk = newInitialSummary.AddElement("DisciplineStatistics", "Demerit");
            string u = "" + dataGridViewX4.Rows[0].Cells[meritList["大過"]].Value;
            string g = "" + dataGridViewX4.Rows[0].Cells[meritList["小過"]].Value;
            string p = "" + dataGridViewX4.Rows[0].Cells[meritList["警告"]].Value;

            int demeritInt;
            if (u == null) //如果是空值
            {
                kkk.SetAttribute("A", "0");
            }
            else if (int.TryParse(u, out demeritInt)) //如果是數字
            {
                kkk.SetAttribute("A", u);
            }

            if (g == null) //如果是空值
            {
                kkk.SetAttribute("B", "0");
            }
            else if (int.TryParse(g, out demeritInt)) //如果是數字
            {
                kkk.SetAttribute("B", g);
            }

            if (p == null) //如果是空值
            {
                kkk.SetAttribute("C", "0");
            }
            else if (int.TryParse(p, out demeritInt)) //如果是數字
            {
                kkk.SetAttribute("C", p);
            }
            #endregion

            #endregion

            #region 取得要計算的資料
            int SaveSchoolYear;
            int SaveSemester;
            string SaveRefStudent;

            if (Mode == "NEW")
            {
                SaveSchoolYear = int.Parse(cbSchoolYear.Text);
                SaveSemester = int.Parse(cbSemester.Text);
                SaveRefStudent = _RefStudentID;
            }
            else
            {
                SaveSchoolYear = _editorRecord.SchoolYear;
                SaveSemester = _editorRecord.Semester;
                SaveRefStudent = _editorRecord.RefStudentID;
            }
            #endregion

            ChangeToCDS CDs = new ChangeToCDS(SaveRefStudent, SaveSchoolYear, SaveSemester, newInitialSummary.BaseElement);

            //看傳出來的Xml正確嗎?

            return CDs.GetXmlElement();
        }

        private void ReflashEffortList()
        {
            #region 努力程度對照表
            EffortList.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["努力程度對照表"];
            if (!string.IsNullOrEmpty(cd["xml"]))
            {
                XmlElement element = XmlHelper.LoadXml(cd["xml"]);

                foreach (XmlElement each in element.SelectNodes("Effort"))
                {
                    EffortList.Add(each.GetAttribute("Code"), each.GetAttribute("Name"));
                }
            }
            #endregion
        }

        private void ReflashDic()
        {
            #region 日常行為表現
            dic.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];
            XmlElement node = XmlHelper.LoadXml(cd["DailyBehavior"]);
            foreach (XmlElement item in node.SelectNodes("PerformanceDegree/Mapping"))
            {
                dic.Add(item.GetAttribute("Degree"), item.GetAttribute("Desc"));
            }
            #endregion
        }

        private void ReflashMorality()
        {
            #region 日常生活表現具體建議使用
            DSResponse dsrsp = Config.GetMoralCommentCodeList();
            Morality.Clear();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Morality"))
            {
                Morality.Add(var.GetAttribute("Code"), var.GetAttribute("Comment"));
            }
            #endregion
        }

        private void CheckSaveButtonEnabled()
        {
            this.btnSave.Enabled = !this.inputErrors.ContainsValue(false);
        }

        private void dgvDailyBehavior_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            #region 日常行為表現資料替換
            dgvDailyBehavior.EndEdit();

            string score = "" + dgvDailyBehavior.CurrentCell.Value;
            bool isMatched = false;

            if (dic.ContainsKey(score)) //如果資料存在key
            {
                dgvDailyBehavior.CurrentCell.Value = dic[score];
                isMatched = true;
            }
            else if (dic.ContainsValue(score)) //如果資料存在value
            {
                dgvDailyBehavior.CurrentCell.Value = score;
                isMatched = true;
            }
            else
            {
                isMatched = false;
            }

            if (string.IsNullOrEmpty(score))
                isMatched = true;

            if (!isMatched)
                dgvDailyBehavior.CurrentCell.Style.BackColor = Color.Pink;
            else
                dgvDailyBehavior.CurrentCell.Style.BackColor = Color.White;

            if (!inputErrors.ContainsKey(dgvDailyBehavior.CurrentCell))
                inputErrors.Add(dgvDailyBehavior.CurrentCell, true);

            inputErrors[dgvDailyBehavior.CurrentCell] = isMatched;
            this.CheckSaveButtonEnabled();

            dgvDailyBehavior.BeginEdit(false);
            #endregion
        }

        private void dgvGroupActivity_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            #region 努力程度資料替換
            dgvGroupActivity.EndEdit();

            bool isMatched = false;
            string score = "" + dgvGroupActivity.CurrentCell.Value;

            if (dgvGroupActivity.CurrentCell.OwningColumn.HeaderText == "努力程度")
            {
                if (EffortList.ContainsKey(score))
                {
                    dgvGroupActivity.CurrentCell.Value = EffortList[score];
                    isMatched = true;
                }
                else if (EffortList.ContainsValue(score)) //如果資料存在value
                {
                    dgvGroupActivity.CurrentCell.Value = score;
                    isMatched = true;
                }
                else
                {
                    isMatched = false;
                }

                if (string.IsNullOrEmpty(score))
                    isMatched = true;

                if (!isMatched)
                    dgvGroupActivity.CurrentCell.Style.BackColor = Color.Pink;
                else
                    dgvGroupActivity.CurrentCell.Style.BackColor = Color.White;

                inputErrors[dgvGroupActivity.CurrentCell] = isMatched;
                this.CheckSaveButtonEnabled();
            }
            dgvGroupActivity.BeginEdit(false);
            #endregion
        }

        public static List<string> GetAbsenceItems()
        {
            #region 取得假別項目

            List<string> list = new List<string>();

            foreach (JHAbsenceMappingInfo element in JHAbsenceMapping.SelectAll())
            {
                list.Add(element.Name);
            }
            return list;
            #endregion
        }

        public static List<string> GetPeriodTypeItems()
        {
            #region 取得節次類型

            List<string> list = new List<string>();

            foreach (JHPeriodMappingInfo element in JHPeriodMapping.SelectAll())
            {
                string type = element.Type;

                if (!list.Contains(type))
                    list.Add(type);
            }
            return list;
            #endregion
        }

        public static List<string> GetMeritTypes()
        {
            #region 取得獎懲清單

            List<string> list = new List<string>();

            list.Add("大功");
            list.Add("小功");
            list.Add("嘉獎");
            list.Add("大過");
            list.Add("小過");
            list.Add("警告");

            return list;

            #endregion
        }

        /// <summary>
        /// 資料檢查
        /// </summary>
        private bool CheckErrorText()
        {
            #region CheckError
            foreach (DataGridViewRow each in dataGridViewX3.Rows)
            {
                foreach (DataGridViewCell each2 in each.Cells)
                {
                    if (each2.ErrorText != "")
                    {
                        return false;
                    }
                }
            }

            foreach (DataGridViewRow each in dataGridViewX4.Rows)
            {
                foreach (DataGridViewCell each2 in each.Cells)
                {
                    if (each2.ErrorText != "")
                    {
                        return false;
                    }
                }
            }
            return true;
            #endregion
        }

        private void dgvDailyLifeRecommend_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            #region 日常生活表現具體建議

            if (dgvDailyLifeRecommend.Rows.Count > 0)
            {
                string daliy = "";
                List<string> listNow = new List<string>();

                if (("" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value).ToString().Contains(','))
                {
                    listNow.AddRange(("" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value).ToString().Split(','));
                }
                else
                {
                    listNow.Add("" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value);
                }

                foreach (string each in listNow)
                {
                    if (daliy == "") //如果是空的
                    {
                        if (Morality.ContainsKey(each))
                        {
                            daliy += Morality[each];
                        }
                        else
                        {
                            daliy += each;
                        }
                    }
                    else //如果不是空的
                    {
                        if (Morality.ContainsKey(each))
                        {
                            daliy += "," + Morality[each];
                        }
                        else
                        {
                            daliy += "," + each;
                        }
                    }
                }
                dgvDailyLifeRecommend.Rows[0].Cells[0].Value = daliy;
            }

            #endregion
        }

        private void dgvDailyLifeRecommend_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgvDailyLifeRecommend.EndEdit();
            dgvDailyLifeRecommend.BeginEdit(false);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void dataGridViewX3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CellValueCheng(dataGridViewX3.Rows[e.RowIndex].Cells[e.ColumnIndex]);
        }

        private void dataGridViewX4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CellValueCheng(dataGridViewX4.Rows[e.RowIndex].Cells[e.ColumnIndex]);
        }

        private void CellValueCheng(DataGridViewCell _cell)
        {
            #region 檢查機制
            string cheng = "" + _cell.Value;
            if (cheng != "")
            {
                int indexNow;
                if (!int.TryParse(cheng, out indexNow))
                {
                    _cell.ErrorText = "資料必須為數字";

                }
                else
                {
                    _cell.ErrorText = "";
                }
            }
            #endregion
        }
    }
}
