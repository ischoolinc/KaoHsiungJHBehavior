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
using K12.Data;
using Framework;
using FISCA.DSAUtil;
using System.Xml;
using FISCA.LogAgent;
using Framework.Feature;

namespace KaoHsiung.DailyLife.ClassDailyLife
{
    public partial class ClassScore : BaseForm
    {

        private BackgroundWorker BGW = new BackgroundWorker();
        private ChangeListener DataListener { get; set; }
        private MoralScoreList MSList; //學生MoralScore清單
        private bool DataGridViewDataInChange = false; //資料是否更動檢查

        private Dictionary<string, int> ColumnIndex = new Dictionary<string, int>(); //Column定位資料
        private Dictionary<string, List<string>> DicSetup = new Dictionary<string, List<string>>();

        private bool CheckBackWorker = false;

        private int SchoolYear;
        private int Semester;

        private Dictionary<string, string> Performance = new Dictionary<string, string>(); //表現程度
        private Dictionary<string, string> EffortList = new Dictionary<string, string>();  //努力程度代碼
        private Dictionary<string, string> Morality = new Dictionary<string, string>(); //日常生活表現具體建議代碼

        public ClassScore()
        {
            InitializeComponent();
        }

        private void ClassDLScoreForm_Load(object sender, EventArgs e)
        {
            ReflashPerformance(); //表現程度
            ReflashEffortList(); //努力程度
            ReflashMorality(); //日常行為

            #region 學年度/學期
            cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear) - 2);
            cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear) - 1);
            cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear));
            cboSchoolYear.Items.Add(int.Parse(School.DefaultSchoolYear) + 1);
            cboSchoolYear.SelectedIndex = 2;
            SchoolYear = int.Parse(School.DefaultSchoolYear);

            cboSemester.Items.Add(1);
            cboSemester.Items.Add(2);
            cboSemester.SelectedIndex = School.DefaultSemester == "1" ? 0 : 1;
            Semester = School.DefaultSemester == "1" ? 1 : 2;
            #endregion

            #region Load

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dataGridViewX1));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            this.Text = "資料載入中,請稍後...";
            this.Enabled = false;

            GetConfigData();

            if (cboPrefs.Items.Count > 0)
                cboPrefs.SelectedIndex = 0;

            cboPrefs.SelectedIndexChanged += new EventHandler(cboPrefs_SelectedIndexChanged);
            cboSchoolYear.SelectedIndexChanged += new EventHandler(cboSchoolYear_SelectedIndexChanged);
            cboSemester.SelectedIndexChanged += new EventHandler(cboSemester_SelectedIndexChanged);

            BGW.RunWorkerAsync();
            #endregion

            dataGridViewX1.Focus();
            // 設定全形半形 (全部欄位)
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dataGridViewX1);
        }

        /// <summary>
        /// 背景模式開始
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            MSList = new MoralScoreList(SchoolYear, Semester);
        }

        /// <summary>
        /// 背景模式完成
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (CheckBackWorker)
            {
                BGW.RunWorkerAsync();
                CheckBackWorker = false;
            }

            this.Text = "評等輸入";
            this.Enabled = true;

            ChangeData();
        }

        /// <summary>
        /// 取得設定值,並且顯示於畫面上
        /// </summary>
        private void GetConfigData()
        {
            DicSetup.Clear();

            #region 取得設定值,並且顯示於畫面上
            K12.Data.Configuration.ConfigData cd = School.Configuration["DLBehaviorConfig"];

            if (cd.Contains("DailyBehavior") && cd["DailyBehavior"] != "")
            {
                XmlElement xml = DSXmlHelper.LoadXml(cd["DailyBehavior"]);
                cboPrefs.Items.Add(xml.GetAttribute("Name"));
                List<string> list = new List<string>();
                foreach (XmlNode each in xml.SelectNodes("Item"))
                {
                    XmlElement eachXml = each as XmlElement;
                    list.Add(eachXml.GetAttribute("Name"));
                }
                DicSetup.Add("DailyBehavior", list);
            }

            if (cd.Contains("GroupActivity") && cd["GroupActivity"] != "")
            {
                XmlElement xml = DSXmlHelper.LoadXml(cd["GroupActivity"]);
                cboPrefs.Items.Add(xml.GetAttribute("Name"));
                List<string> list = new List<string>();
                foreach (XmlNode each in xml.SelectNodes("Item"))
                {
                    XmlElement eachXml = each as XmlElement;
                    list.Add(eachXml.GetAttribute("Name") + "：努力程度");
                    list.Add(eachXml.GetAttribute("Name") + "：文字描述");
                }
                DicSetup.Add("GroupActivity", list);
            }

            if (cd.Contains("PublicService") && cd["PublicService"] != "")
            {
                XmlElement xml = DSXmlHelper.LoadXml(cd["PublicService"]);
                cboPrefs.Items.Add(xml.GetAttribute("Name"));
                List<string> list = new List<string>();
                foreach (XmlNode each in xml.SelectNodes("Item"))
                {
                    XmlElement eachXml = each as XmlElement;
                    list.Add(eachXml.GetAttribute("Name"));
                }
                DicSetup.Add("PublicService", list);
            }

            if (cd.Contains("SchoolSpecial") && cd["SchoolSpecial"] != "")
            {
                XmlElement xml = DSXmlHelper.LoadXml(cd["SchoolSpecial"]);
                cboPrefs.Items.Add(xml.GetAttribute("Name"));
                List<string> list = new List<string>();
                foreach (XmlNode each in xml.SelectNodes("Item"))
                {
                    XmlElement eachXml = each as XmlElement;
                    list.Add(eachXml.GetAttribute("Name"));
                }
                DicSetup.Add("SchoolSpecial", list);
            }

            if (cd.Contains("DailyLifeRecommend") && cd["DailyLifeRecommend"] != "")
            {
                XmlElement xml = DSXmlHelper.LoadXml(cd["DailyLifeRecommend"]);
                cboPrefs.Items.Add(xml.GetAttribute("Name"));

                List<string> list = new List<string>();
                list.Add(xml.GetAttribute("Name"));

                DicSetup.Add("DailyLifeRecommend", list);
            }
            #endregion

        }

        /// <summary>
        /// 如果DataGridView內容變更了
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            DataGridViewDataInChange = true;
        }

        private void ChangeMessage()
        {
            #region 確認訊息
            if (DataGridViewDataInChange)
            {
                DialogResult dr = Framework.MsgBox.Show("您資料已變更,是否要儲存資料?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button1);
                if (dr == DialogResult.Yes)
                {
                    if (btnSave.Enabled)
                    {
                        btnSave_Click(null, null);
                    }
                    else
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("資料錯誤,無法儲存!!");
                        dataGridViewX1.Focus();
                        return;
                    }
                }

                DataGridViewDataInChange = false;
            }
            #endregion
        }

        private void ChangeData()
        {
            #region 更新頁面資料
            DataListener.SuspendListen(); //終止變更判斷

            SaveMoralScoreRecordList.Clear();

            ColumnCheng();
            DataCheng();

            DataListener.Reset();
            DataListener.ResumeListen();
            #endregion
        }

        /// <summary>
        /// 更新Column
        /// </summary>
        private void ColumnCheng()
        {
            #region 更新Column
            dataGridViewX1.Rows.Clear();
            dataGridViewX1.Columns.Clear();
            ColumnIndex.Clear();

            SetColumnName("ID", 0);
            SetColumnName("班級", 65);
            SetColumnName("座號", 65);
            SetColumnName("姓名", 65);


            if (cboPrefs.SelectedIndex == 0) //DailyBehavior
            {
                foreach (string each in DicSetup["DailyBehavior"])
                {
                    SetColumnNameLock(each, 90);
                }
            }
            else if (cboPrefs.SelectedIndex == 1) //GroupActivity
            {
                foreach (string each in DicSetup["GroupActivity"])
                {
                    SetColumnNameLock(each, 100);
                }
            }
            else if (cboPrefs.SelectedIndex == 2) //PublicService
            {
                foreach (string each in DicSetup["PublicService"])
                {
                    SetColumnNameLock(each, 130);
                }
            }
            else if (cboPrefs.SelectedIndex == 3) //SchoolSpecial
            {
                foreach (string each in DicSetup["SchoolSpecial"])
                {
                    SetColumnNameLock(each, 200);
                }
            }
            else if (cboPrefs.SelectedIndex == 4) //DailyLifeRecommend
            {
                foreach (string each in DicSetup["DailyLifeRecommend"])
                {
                    SetColumnNameLock(each, 400);
                }
            }
            #endregion
        }

        /// <summary>
        /// 更新畫面資料
        /// </summary>
        private void DataCheng()
        {
            #region 更新畫面資料
            foreach (JHStudentRecord student in MSList.StudentRe)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);

                row.Cells[ColumnIndex["ID"]].Value = student.ID;
                row.Cells[ColumnIndex["班級"]].Value = student.Class.Name;
                row.Cells[ColumnIndex["座號"]].Value = student.SeatNo;
                row.Cells[ColumnIndex["姓名"]].Value = student.Name;

                JHMoralScoreRecord StudentSR = MSList.GetMoralScore(student.ID);

                if (StudentSR != null)
                {
                    if (cboPrefs.SelectedIndex == 0) //DailyBehavior
                    {
                        #region DailyBehavior
                        XmlElement StudentDBXml = StudentSR.TextScore.SelectSingleNode("DailyBehavior") as XmlElement;

                        if (StudentDBXml != null)
                        {
                            Dictionary<string, string> dic = new Dictionary<string, string>();

                            foreach (XmlNode SelectEach in StudentDBXml.SelectNodes("Item"))
                            {
                                XmlElement xml = SelectEach as XmlElement;
                                dic.Add(xml.GetAttribute("Name"), xml.GetAttribute("Degree"));
                            }

                            foreach (string db in DicSetup["DailyBehavior"])
                            {
                                if (dic.ContainsKey(db))
                                {
                                    row.Cells[ColumnIndex[db]].Value = dic[db];
                                }
                            }
                        }
                        #endregion
                    }
                    else if (cboPrefs.SelectedIndex == 1) //GroupActivity
                    {
                        #region GroupActivity
                        XmlElement StudentDBXml = StudentSR.TextScore.SelectSingleNode("GroupActivity") as XmlElement;
                        if (StudentDBXml != null)
                        {
                            Dictionary<string, string> dic = new Dictionary<string, string>();
                            foreach (XmlNode SelectEach in StudentDBXml.SelectNodes("Item"))
                            {
                                XmlElement xml = SelectEach as XmlElement;
                                dic.Add(xml.GetAttribute("Name") + "：努力程度", xml.GetAttribute("Degree"));
                                dic.Add(xml.GetAttribute("Name") + "：文字描述", xml.GetAttribute("Description"));
                            }

                            foreach (string db in DicSetup["GroupActivity"])
                            {
                                if (dic.ContainsKey(db))
                                {
                                    row.Cells[ColumnIndex[db]].Value = dic[db];
                                }
                            }
                        }
                        #endregion
                    }
                    else if (cboPrefs.SelectedIndex == 2) //PublicService
                    {
                        #region PublicService
                        XmlElement StudentDBXml = StudentSR.TextScore.SelectSingleNode("PublicService") as XmlElement;
                        if (StudentDBXml != null)
                        {
                            Dictionary<string, string> dic = new Dictionary<string, string>();
                            foreach (XmlNode SelectEach in StudentDBXml.SelectNodes("Item"))
                            {
                                XmlElement xml = SelectEach as XmlElement;
                                dic.Add(xml.GetAttribute("Name"), xml.GetAttribute("Description"));
                            }

                            foreach (string db in DicSetup["PublicService"])
                            {
                                if (dic.ContainsKey(db)) //如果取得的資料,不包含在設定檔內
                                {
                                    row.Cells[ColumnIndex[db]].Value = dic[db];
                                }
                            }
                        }
                        #endregion
                    }
                    else if (cboPrefs.SelectedIndex == 3) //SchoolSpecial
                    {
                        #region SchoolSpecial
                        XmlElement StudentDBXml = StudentSR.TextScore.SelectSingleNode("SchoolSpecial") as XmlElement;
                        if (StudentDBXml != null)
                        {
                            Dictionary<string, string> dic = new Dictionary<string, string>();
                            foreach (XmlNode SelectEach in StudentDBXml.SelectNodes("Item"))
                            {
                                XmlElement xml = SelectEach as XmlElement;
                                dic.Add(xml.GetAttribute("Name"), xml.GetAttribute("Description"));
                            }

                            foreach (string db in DicSetup["SchoolSpecial"])
                            {
                                if (dic.ContainsKey(db))
                                {
                                    row.Cells[ColumnIndex[db]].Value = dic[db];
                                }
                            }
                        }
                        #endregion
                    }
                    else if (cboPrefs.SelectedIndex == 4) //DailyLifeRecommend
                    {
                        #region DailyLifeRecommend
                        XmlElement StudentDBXml = StudentSR.TextScore.SelectSingleNode("DailyLifeRecommend") as XmlElement;
                        if (StudentDBXml != null)
                        {
                            foreach (string db in DicSetup["DailyLifeRecommend"])
                            {
                                row.Cells[ColumnIndex[db]].Value = StudentDBXml.GetAttribute("Description");
                            }
                        }
                        #endregion
                    }
                }

                dataGridViewX1.Rows.Add(row);
            }
            #endregion
        }

        /// <summary>
        /// 儲存畫面內容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            this.Enabled = false;

            SaveData();
            this.Enabled = true;
            DataGridViewDataInChange = false;
            FISCA.Presentation.Controls.MsgBox.Show("儲存完成");
        }

        private void SaveData()
        {
            #region Save
            //分為新增和更新2種資料內容
            List<JHMoralScoreRecord> UPdataMSR = new List<JHMoralScoreRecord>();
            List<JHMoralScoreRecord> UPInsertMSR = new List<JHMoralScoreRecord>();
            //取得設定值
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];
            //更新學生資料
            MSList = new MoralScoreList(SchoolYear, Semester);

            foreach (DataGridViewRow each in dataGridViewX1.Rows)
            {
                #region 驗證
                //取得學生ID
                string RowStudentId = "" + each.Cells[0].Value;

                if (!SaveMoralScoreRecordList.Contains(RowStudentId))
                {
                    continue;
                }

                //取得學生JHMoralScoreRecord
                JHMoralScoreRecord StudentMoralScoreRecord = MSList.GetMoralScore(RowStudentId);

                //如果沒有則新增
                if (StudentMoralScoreRecord == null)
                {
                    StudentMoralScoreRecord = new JHMoralScoreRecord();
                    StudentMoralScoreRecord.RefStudentID = RowStudentId;
                    StudentMoralScoreRecord.SchoolYear = SchoolYear;
                    StudentMoralScoreRecord.Semester = Semester;
                }

                XmlElement textscore = null;
                DSXmlHelper hlptextscore = null;

                //如果TextScore欄位空值
                if (StudentMoralScoreRecord.TextScore == null)
                    StudentMoralScoreRecord.TextScore = DSXmlHelper.LoadXml("<TextScore/>");

                textscore = StudentMoralScoreRecord.TextScore;

                //當做DSXmlHelper物件操作
                hlptextscore = new DSXmlHelper(textscore);
                #endregion

                if (cboPrefs.SelectedIndex == 0)
                {
                    #region DailyBehavior
                    if (cd.Contains("DailyBehavior"))
                    {

                        //取得Element
                        if (hlptextscore.GetElement("DailyBehavior") == null)
                        {
                            hlptextscore.AddElement("DailyBehavior");
                        }
                        //刪除原Element內容
                        hlptextscore.GetElement("DailyBehavior").RemoveAll();
                        //取得DailyBehavior的設定內容
                        XmlElement node = DSXmlHelper.LoadXml(cd["DailyBehavior"]);
                        //在hlptextscore內依設定值建立架構
                        hlptextscore.GetElement("DailyBehavior").SetAttribute("Name", node.GetAttribute("Name"));
                        foreach (XmlElement item in node.SelectNodes("Item"))
                        {
                            string name = item.GetAttribute("Name");
                            XmlElement anode = hlptextscore.AddElement("DailyBehavior", "Item");
                            anode.SetAttribute("Name", name);
                            anode.SetAttribute("Index", item.GetAttribute("Index"));
                        }

                        //取出DailyBehavior內容
                        XmlElement knode = hlptextscore.GetElement("DailyBehavior");
                        //填入Degree的值

                        foreach (XmlElement Eachnode in knode.SelectNodes("Item"))
                        {
                            string ColumnName = Eachnode.GetAttribute("Name");
                            if (ColumnIndex.ContainsKey(ColumnName)) //是否有這個欄位
                            {
                                int ColumnXXX = ColumnIndex[ColumnName];
                                Eachnode.SetAttribute("Degree", "" + each.Cells[ColumnXXX].Value);
                            }
                        }
                    }
                    #endregion
                }
                else if (cboPrefs.SelectedIndex == 1)
                {
                    #region GroupActivity
                    if (cd.Contains("GroupActivity"))
                    {
                        //取得Element
                        if (hlptextscore.GetElement("GroupActivity") == null)
                        {
                            hlptextscore.AddElement("GroupActivity");
                        }

                        hlptextscore.GetElement("GroupActivity").RemoveAll();

                        XmlElement node = DSXmlHelper.LoadXml(cd["GroupActivity"]);
                        hlptextscore.GetElement("GroupActivity").SetAttribute("Name", node.GetAttribute("Name"));

                        foreach (XmlElement item in node.SelectNodes("Item"))
                        {
                            XmlElement anode = hlptextscore.AddElement("GroupActivity", "Item");
                            anode.SetAttribute("Name", item.GetAttribute("Name"));
                        }

                        XmlElement gnode = hlptextscore.GetElement("GroupActivity");


                        foreach (XmlElement EachNode in gnode.SelectNodes("Item"))
                        {
                            string ColumnName = EachNode.GetAttribute("Name") + "：努力程度";
                            if (ColumnIndex.ContainsKey(ColumnName))
                            {
                                int ColumnItemIndex = ColumnIndex[ColumnName]; //取得Index
                                EachNode.SetAttribute("Degree", "" + each.Cells[ColumnItemIndex].Value);
                                EachNode.SetAttribute("Description", "" + each.Cells[ColumnItemIndex + 1].Value);
                            }
                        }
                    }
                    #endregion
                }
                else if (cboPrefs.SelectedIndex == 2)
                {
                    #region PublicService
                    if (cd.Contains("PublicService"))
                    {
                        if (hlptextscore.GetElement("PublicService") == null)
                            hlptextscore.AddElement("PublicService");

                        hlptextscore.GetElement("PublicService").RemoveAll();

                        XmlElement node = DSXmlHelper.LoadXml(cd["PublicService"]);
                        hlptextscore.GetElement("PublicService").SetAttribute("Name", node.GetAttribute("Name"));

                        foreach (XmlElement item in node.SelectNodes("Item"))
                        {
                            XmlElement anode = hlptextscore.AddElement("PublicService", "Item");
                            anode.SetAttribute("Name", item.GetAttribute("Name"));
                        }

                        XmlElement pnode = hlptextscore.GetElement("PublicService");

                        foreach (XmlElement EachNode in pnode.SelectNodes("Item"))
                        {
                            string ColumnName = EachNode.GetAttribute("Name");
                            if (ColumnIndex.ContainsKey(ColumnName))
                            {
                                int ColumnItemIndex = ColumnIndex[ColumnName];
                                EachNode.SetAttribute("Description", "" + each.Cells[ColumnItemIndex].Value);
                            }
                        }
                    }
                    #endregion
                }
                else if (cboPrefs.SelectedIndex == 3)
                {
                    #region SchoolSpecial
                    if (cd.Contains("SchoolSpecial"))
                    {
                        if (hlptextscore.GetElement("SchoolSpecial") == null)
                            hlptextscore.AddElement("SchoolSpecial");

                        hlptextscore.GetElement("SchoolSpecial").RemoveAll();

                        XmlElement node = DSXmlHelper.LoadXml(cd["SchoolSpecial"]);
                        hlptextscore.GetElement("SchoolSpecial").SetAttribute("Name", node.GetAttribute("Name"));

                        foreach (XmlElement item in node.SelectNodes("Item"))
                        {
                            XmlElement anode = hlptextscore.AddElement("SchoolSpecial", "Item");
                            anode.SetAttribute("Name", item.GetAttribute("Name"));
                        }

                        XmlElement qnode = hlptextscore.GetElement("SchoolSpecial");

                        foreach (XmlElement EachNode in qnode.SelectNodes("Item"))
                        {
                            string ColumnName = EachNode.GetAttribute("Name");
                            if (ColumnIndex.ContainsKey(ColumnName))
                            {
                                int ColumnItemIndex = ColumnIndex[ColumnName];
                                EachNode.SetAttribute("Description", "" + each.Cells[ColumnItemIndex].Value);
                            }
                        }
                    }
                    #endregion
                }
                else if (cboPrefs.SelectedIndex == 4)
                {
                    #region DailyLifeRecommend
                    if (cd.Contains("DailyLifeRecommend"))
                    {
                        //如果學生沒有此欄位
                        if (hlptextscore.GetElement("DailyLifeRecommend") == null)
                            hlptextscore.AddElement("DailyLifeRecommend");

                        //清空此Element內所有元素
                        hlptextscore.GetElement("DailyLifeRecommend").RemoveAll();

                        //讀取設定值內容
                        XmlElement node = DSXmlHelper.LoadXml(cd["DailyLifeRecommend"]);

                        //設定此Element的名稱
                        hlptextscore.GetElement("DailyLifeRecommend").SetAttribute("Name", node.GetAttribute("Name"));

                        XmlElement qnode = hlptextscore.GetElement("DailyLifeRecommend");
                        string ColumnName = qnode.GetAttribute("Name");
                        if (ColumnIndex.ContainsKey(ColumnName))
                        {
                            int ColumnItemIndex = ColumnIndex[ColumnName];
                            hlptextscore.GetElement("DailyLifeRecommend").SetAttribute("Description", "" + each.Cells[ColumnItemIndex].Value);
                        }
                    }

                    #endregion
                }

                if (IsAddRequired(StudentMoralScoreRecord))
                {
                    UPInsertMSR.Add(StudentMoralScoreRecord);
                }
                else
                {
                    UPdataMSR.Add(StudentMoralScoreRecord);
                }
            }

            try
            {
                JHMoralScore.Insert(UPInsertMSR);
                JHMoralScore.Update(UPdataMSR);
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("儲存發生錯誤");
                throw ex;
            }

            #region Log
            if (UPInsertMSR.Count != 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("由「評等輸入」功能，將學生「日常生活表現資料」進行批次新增。");
                sb.AppendLine("詳細資料：\n學年度「" + SchoolYear.ToString() + "」學期「" + Semester.ToString() + "」。");
                foreach (JHMoralScoreRecord sbEach in UPInsertMSR)
                {
                    StringBuilder setbu = SetLog("新增資料", sbEach);
                    sb.Append(setbu.ToString());
                }

                ApplicationLog.Log("日常生活表現模組.評等輸入", "新增日常生活表現資料", sb.ToString());

            }

            if (UPdataMSR.Count != 0)
            {
                StringBuilder sc = new StringBuilder();

                sc.AppendLine("由「評等輸入」功能，將學生「日常生活表現資料」進行批次修改。");
                sc.AppendLine("詳細資料：\n學年度「" + SchoolYear.ToString() + "」學期「" + Semester.ToString() + "」。");

                foreach (JHMoralScoreRecord sbEach in UPdataMSR)
                {
                    StringBuilder setbu = SetLog("更新資料", sbEach);
                    sc.Append(setbu.ToString());
                }

                ApplicationLog.Log("日常生活表現模組.評等輸入", "修改日常生活表現資料", sc.ToString());

            }
            #endregion

            #endregion
        }

        private StringBuilder SetLog(string IsUpdataOrInsert, JHMoralScoreRecord sbEach)
        {
            StringBuilder sc = new StringBuilder();
            sc.Append("班級「" + sbEach.Student.Class.Name + "」");
            if (sbEach.Student.SeatNo.HasValue)
                sc.Append("座號「" + sbEach.Student.SeatNo.Value + "」");
            else
                sc.Append("座號「」");

            sc.AppendLine("學生「" + sbEach.Student.Name + "」");

            if (cboPrefs.SelectedIndex == 0)
            {
                XmlElement xml1 = (XmlElement)sbEach.TextScore.SelectSingleNode("DailyBehavior");
                if (xml1 != null)
                {
                    sc.Append(IsUpdataOrInsert + "「" + xml1.GetAttribute("Name") + "」");
                    foreach (XmlElement xmleach in xml1.SelectNodes("Item"))
                    {
                        sc.Append(xmleach.GetAttribute("Name") + "「" + xmleach.GetAttribute("Degree") + "」");
                    }
                    sc.AppendLine(); //換行
                }

            }
            else if (cboPrefs.SelectedIndex == 1)
            {
                XmlElement xml1 = (XmlElement)sbEach.TextScore.SelectSingleNode("GroupActivity");
                if (xml1 != null)
                {
                    sc.Append(IsUpdataOrInsert + "「" + xml1.GetAttribute("Name") + "」");
                    foreach (XmlElement xmleach in xml1.SelectNodes("Item"))
                    {
                        sc.Append(xmleach.GetAttribute("Name") + "：" + "努力程度" + "「" + xmleach.GetAttribute("Degree") + "」" + "文字描述「" + xmleach.GetAttribute("Description") + "」");
                    }
                    sc.AppendLine(); //換行
                }
            }
            else if (cboPrefs.SelectedIndex == 2)
            {
                XmlElement xml1 = (XmlElement)sbEach.TextScore.SelectSingleNode("PublicService");
                if (xml1 != null)
                {
                    sc.Append(IsUpdataOrInsert + "「" + xml1.GetAttribute("Name") + "」");
                    foreach (XmlElement xmleach in xml1.SelectNodes("Item"))
                    {
                        sc.Append(xmleach.GetAttribute("Name") + "：" + "文字描述" + "「" + xmleach.GetAttribute("Description") + "」");
                    }
                    sc.AppendLine(); //換行
                }
            }
            else if (cboPrefs.SelectedIndex == 3)
            {
                XmlElement xml1 = (XmlElement)sbEach.TextScore.SelectSingleNode("SchoolSpecial");
                if (xml1 != null)
                {
                    sc.Append(IsUpdataOrInsert + "「" + xml1.GetAttribute("Name") + "」");
                    foreach (XmlElement xmleach in xml1.SelectNodes("Item"))
                    {
                        sc.Append(xmleach.GetAttribute("Name") + "：" + "文字描述" + "「" + xmleach.GetAttribute("Description") + "」");
                    }
                    sc.AppendLine(); //換行
                }
            }
            else if (cboPrefs.SelectedIndex == 4)
            {
                XmlElement xml1 = (XmlElement)sbEach.TextScore.SelectSingleNode("DailyLifeRecommend");
                if (xml1 != null)
                {
                    sc.Append(IsUpdataOrInsert + "「" + xml1.GetAttribute("Name") + "」");
                    sc.AppendLine("文字描述「" + xml1.GetAttribute("Description") + "」");
                }
            }

            return sc;
        }

        private static bool IsAddRequired(JHMoralScoreRecord editor)
        {
            return string.IsNullOrEmpty(editor.ID);
        }

        #region 對Column操作
        /// <summary>
        /// 新增Column
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnName(string x, int y)
        {
            int NUM = dataGridViewX1.Columns.Add(x, x);
            dataGridViewX1.Columns[NUM].Width = y;
            ColumnIndex.Add(x, NUM);
            SetColumnStyle(NUM);
        }

        /// <summary>
        /// 新增Column,不鎖定
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnNameLock(string x, int y)
        {
            int NUM = dataGridViewX1.Columns.Add(x, x);
            dataGridViewX1.Columns[NUM].Width = y;
            ColumnIndex.Add(x, NUM);
        }

        /// <summary>
        /// 設定預設Column樣式
        /// </summary>
        /// <param name="x"></param>
        private void SetColumnStyle(int x)
        {
            dataGridViewX1.Columns[x].ReadOnly = true;
            dataGridViewX1.Columns[x].DefaultCellStyle.BackColor = Color.LightCyan;

            if (x == 0)
            {
                dataGridViewX1.Columns[x].Visible = false; //如果是ID要隱藏
            }
        }
        #endregion

        #region 切換資料操作(事件)
        private void cboPrefs_Enter(object sender, EventArgs e)
        {
            ChangeMessage();
        }

        private void cboSchoolYear_Enter(object sender, EventArgs e)
        {
            ChangeMessage();
        }

        private void cboSemester_Enter(object sender, EventArgs e)
        {
            ChangeMessage();
        }
        #endregion

        #region 更新資料(事件)
        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            SchoolYear = int.Parse(cboSchoolYear.Text);

            if (BGW.IsBusy)
            {
                CheckBackWorker = true;
            }
            else
            {
                BGW.RunWorkerAsync();
            }
            ChangeData();
        }

        private void cboSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            Semester = int.Parse(cboSemester.Text);

            if (BGW.IsBusy)
            {
                CheckBackWorker = true;
            }
            else
            {
                BGW.RunWorkerAsync();
            }
            ChangeData();
        }

        private void cboPrefs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (BGW.IsBusy)
            {
                CheckBackWorker = true;
            }
            else
            {
                BGW.RunWorkerAsync();
            }
            ChangeData();
        }
        #endregion

        /// <summary>
        /// 離開
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExit_Click(object sender, EventArgs e)
        {
            if (DataGridViewDataInChange)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("資料已被修改,請確認是否要關閉?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    this.Close();
                }
            }
            else
            {
                this.Close();
            }
        }

        private void ReflashPerformance()
        {
            #region 表現程度對照表
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];
            Performance.Clear();

            if (string.IsNullOrEmpty(cd["DailyBehavior"]))
            {
                Framework.MsgBox.Show("日常生活表現設定檔發現錯誤,請重新設定");
                return;
            }

            XmlElement node = DSXmlHelper.LoadXml(cd["DailyBehavior"]);

            if (node.SelectNodes("PerformanceDegree/Mapping") == null)
            {
                Framework.MsgBox.Show("尚未設定表現程度對照表,將無法自動替換代碼!");
                return;
            }
            else if (node.SelectNodes("PerformanceDegree/Mapping").Count == 0)
            {
                Framework.MsgBox.Show("尚未設定表現程度對照表,將無法自動替換代碼!");
                return;
            }

            foreach (XmlElement item in node.SelectNodes("PerformanceDegree/Mapping"))
            {
                Performance.Add(item.GetAttribute("Degree"), item.GetAttribute("Desc"));
            }
            #endregion
        }

        private void ReflashEffortList()
        {
            #region 努力程度對照表
            EffortList.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["努力程度對照表"];
            if (!string.IsNullOrEmpty(cd["xml"]))
            {
                XmlElement element = DSXmlHelper.LoadXml(cd["xml"]);

                foreach (XmlElement each in element.SelectNodes("Effort"))
                {
                    EffortList.Add(each.GetAttribute("Code"), each.GetAttribute("Name"));
                }
            }
            #endregion
        }

        private void ReflashMorality()
        {
            #region 日常行為表現對照表
            Morality.Clear();
            DSResponse dsrsp = Config.GetMoralCommentCodeList();
            foreach (XmlElement var in dsrsp.GetContent().GetElements("Morality"))
            {
                Morality.Add(var.GetAttribute("Code"), var.GetAttribute("Comment"));
            }
            #endregion
        }

        List<string> SaveMoralScoreRecordList = new List<string>();

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            //如果資料進行變更
            if (!SaveMoralScoreRecordList.Contains("" + dataGridViewX1.CurrentRow.Cells[0].Value))
            {
                SaveMoralScoreRecordList.Add("" + dataGridViewX1.CurrentRow.Cells[0].Value);
            }

            #region 替換代碼
            dataGridViewX1.EndEdit();
            DataGridViewCell CurrCell = dataGridViewX1.CurrentCell;
            string score = "" + CurrCell.Value;

            if (cboPrefs.SelectedIndex == 0)
            {
                #region DailyBehavior
                if (Performance.ContainsKey(score)) //如果資料存在key
                {
                    CurrCell.Value = Performance[score];
                    CurrCell.Style.BackColor = Color.White;
                }
                else if (Performance.ContainsValue(score)) //如果資料存在value
                {
                    CurrCell.Style.BackColor = Color.White;
                }
                else if (score == "")
                {
                    CurrCell.Style.BackColor = Color.White;
                }
                else
                {
                    CurrCell.Style.BackColor = Color.Pink;
                }
                #endregion
            }
            else if (cboPrefs.SelectedIndex == 1) //替換努力程度
            {
                #region GroupActivity
                if (CurrCell.OwningColumn.HeaderText.Contains("努力程度"))
                {

                    if (EffortList.ContainsKey(score))
                    {
                        CurrCell.Value = EffortList[score];
                        CurrCell.Style.BackColor = Color.White;
                    }
                    else if (EffortList.ContainsValue(score))
                    {
                        CurrCell.Style.BackColor = Color.White;
                    }
                    else if (score == "")
                    {
                        CurrCell.Style.BackColor = Color.White;
                    }
                    else
                    {
                        CurrCell.Style.BackColor = Color.Pink;
                    }
                }
                #endregion
            }
            CheckSaveButtonEnabled();
            dataGridViewX1.BeginEdit(false);
            #endregion
        }

        private void dataGridViewX1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            #region 替換代碼(DailyLifeRecommend)
            if (cboPrefs.SelectedIndex == 4)
            {
                DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                //cell.Value;

                string daliy = "";
                if (cell.Value == null)
                    return;
                string NowCell = "" + cell.Value.ToString();
                List<string> listNow = new List<string>();

                if (NowCell.Contains(','))
                {
                    listNow.AddRange(NowCell.Split(','));
                }
                else
                {
                    listNow.Add(NowCell);
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
                cell.Value = daliy;
            }
            #endregion


        }

        /// <summary>
        /// 檢查儲存按鈕是否可以按。當格子裡沒有錯誤的值時才Enabled。
        /// </summary>       
        private void CheckSaveButtonEnabled()
        {
            #region 資料檢查
            int Count = 0;
            foreach (DataGridViewRow eachRow in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell eachCell in eachRow.Cells)
                {
                    if (eachCell.Style.BackColor == Color.Pink)
                    {
                        Count++;
                    }
                }
            }

            if (Count == 0)
            {
                btnSave.Enabled = true;
            }
            else
            {
                btnSave.Enabled = false;
            }
            #endregion
        }
    }
}
