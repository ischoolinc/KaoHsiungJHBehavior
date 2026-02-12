using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Framework;
using System.Xml;
using Framework.Feature;
using FISCA.Presentation.Controls;
using JHSchool;
using JHSchool.Data;
using FISCA.DSAUtil;
using FISCA.LogAgent;
using Framework;

namespace KaoHsiung.DailyLife
{
    public partial class DLScoreEditForm : BaseForm
    {
        //<TextScore>
        //    <DailyBehavior Name="日常行為表現">
        //        <Item Name="愛整潔" Index="....." Degree="3"/>
        //        <Item Name="守秩序" Index="....." Degree="3"/>
        //    </DailyBehavior>

        //    <GroupActivity Name="團體活動表現">
        //        <Item Name="社團活動" Degree="1" Description=".....">
        //        <Item Name="學校活動" Degree="2" Description=".....">
        //    </GroupActivity>

        //    <PublicService Name="公共服務表現">
        //        <Item Name="校內服務" Description=".....">
        //        <Item Name="社區服務" Description=".....">
        //    </PublicService>

        //    <SchoolSpecial Name="校內外時特殊表現">
        //        <Item Name="校外特殊表現" Description=".....">
        //        <Item Name="校內特殊表現" Description=".....">
        //    </SchoolSpecial>

        //    <DailyLifeRecommend Name="日常生活表現具體建議" Description=".....">
        //</TextScore>

        private JHSchool.Data.JHMoralScoreRecord _editorRecord;
        private Dictionary<DataGridViewCell, bool> inputErrors = new Dictionary<DataGridViewCell, bool>();     //用來記錄 日常行為表現 是否有格子輸入不正確的值
        private Dictionary<string, string> Morality = new Dictionary<string, string>(); //日常生活表現具體建議使用
        private Dictionary<string, string> EffortList = new Dictionary<string, string>();  //努力程度代碼
        private Dictionary<string, string> dic = new Dictionary<string, string>();
        private JHSchool.Data.JHStudentRecord _SR;
        private Framework.Security.FeatureAce _permission;

        StringBuilder sb_log = new StringBuilder();

        //是否變更資料
        private ChangeListener DataListener { get; set; }
        bool CheckChange = false;

        private string Mode;

        string _PrimaryKey;

        public DLScoreEditForm(string PrimaryKey, Framework.Security.FeatureAce permission)
        {
            InitializeComponent();

            _permission = permission; //權限

            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dgvDailyBehavior));
            DataListener.Add(new DataGridViewSource(dgvGroupActivity));
            DataListener.Add(new DataGridViewSource(dgvPublicService));
            DataListener.Add(new DataGridViewSource(dgvSchoolSpecial));
            DataListener.Add(new DataGridViewSource(dgvDailyLifeRecommend));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            _PrimaryKey = PrimaryKey;

            _SR = JHStudent.SelectByID(_PrimaryKey); //取得學生

            #region 學年度學期
            string schoolYear = K12.Data.School.DefaultSchoolYear;
            cbSchoolYear.Text = schoolYear;
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 3).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 2).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) - 1).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear)).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 1).ToString());
            cbSchoolYear.Items.Add((int.Parse(schoolYear) + 2).ToString());

            string semester = K12.Data.School.DefaultSemester;
            cbSemester.Text = semester;
            cbSemester.Items.Add("1");
            cbSemester.Items.Add("2");
            #endregion

            this.cbSchoolYear.TextChanged += new System.EventHandler(this.cbSchoolYear_TextChanged);
            this.cbSemester.TextChanged += new System.EventHandler(this.cbSemester_TextChanged);

            JHMoralScoreRecord CheckMSR = JHMoralScore.SelectBySchoolYearAndSemester(_PrimaryKey, int.Parse(schoolYear), int.Parse(semester));
            if (CheckMSR == null)
            {
                Mode = "NEW";
                _editorRecord = new JHMoralScoreRecord();
                SyndLoad();
            }
            else
            {
                Mode = "UPDATA";
                _editorRecord = CheckMSR;

                SyndLoad();

                BindData();
            }

            DataListener.Reset();
            DataListener.ResumeListen();

        }

        public DLScoreEditForm(JHSchool.Data.JHMoralScoreRecord editor, Framework.Security.FeatureAce permission)
        {
            #region 建構子
            InitializeComponent();

            _PrimaryKey = editor.RefStudentID;

            Mode = "UPDATA";

            _permission = permission; //權限

            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dgvDailyBehavior));
            DataListener.Add(new DataGridViewSource(dgvGroupActivity));
            DataListener.Add(new DataGridViewSource(dgvPublicService));
            DataListener.Add(new DataGridViewSource(dgvSchoolSpecial));
            DataListener.Add(new DataGridViewSource(dgvDailyLifeRecommend));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            _editorRecord = editor;

            _SR = JHStudent.SelectByID(_editorRecord.RefStudentID); //取得學生

            cbSchoolYear.Text = _editorRecord.SchoolYear.ToString();
            cbSchoolYear.Enabled = false;
            cbSemester.Text = _editorRecord.Semester.ToString();
            cbSemester.Enabled = false;

            SyndLoad();

            BindData();

            DataListener.Reset();
            DataListener.ResumeListen();

            #endregion

        }

        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            CheckChange = (e.Status == ValueStatus.Dirty);
        }

        /// <summary>
        /// 初始化表格規格
        /// </summary>
        private void SyndLoad()
        {
            dgvDailyBehavior.Rows.Clear();
            dgvGroupActivity.Rows.Clear();
            dgvPublicService.Rows.Clear();
            dgvSchoolSpecial.Rows.Clear();
            dgvDailyLifeRecommend.Rows.Clear();
            dgvDailyLifeRecommend.Rows.Add();

            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];

            if (!string.IsNullOrEmpty(cd["DailyBehavior"]))
            {
                XmlElement dailyBehavior = K12.Data.XmlHelper.LoadXml(cd["DailyBehavior"]);
                tabControl1.Tabs[0].Text = dailyBehavior.GetAttribute("Name");

                foreach (XmlElement item in dailyBehavior.SelectNodes("Item"))
                    dgvDailyBehavior.Rows.Add(item.GetAttribute("Name"), item.GetAttribute("Index"), "");
            }

            if (!string.IsNullOrEmpty(cd["GroupActivity"]))
            {
                XmlElement groupActivity = K12.Data.XmlHelper.LoadXml(cd["GroupActivity"]);
                tabControl1.Tabs[1].Text = groupActivity.GetAttribute("Name");

                foreach (XmlElement item in groupActivity.SelectNodes("Item"))
                    dgvGroupActivity.Rows.Add(item.GetAttribute("Name"), "", "");
            }

            if (!string.IsNullOrEmpty(cd["PublicService"]))
            {
                XmlElement publicService = K12.Data.XmlHelper.LoadXml(cd["PublicService"]);
                tabControl1.Tabs[2].Text = publicService.GetAttribute("Name");

                foreach (XmlElement item in publicService.SelectNodes("Item"))
                    dgvPublicService.Rows.Add(item.GetAttribute("Name"), "");
            }

            if (!string.IsNullOrEmpty(cd["SchoolSpecial"]))
            {
                XmlElement schoolSpecial = K12.Data.XmlHelper.LoadXml(cd["SchoolSpecial"]);
                tabControl1.Tabs[3].Text = schoolSpecial.GetAttribute("Name");

                foreach (XmlElement item in schoolSpecial.SelectNodes("Item"))
                    dgvSchoolSpecial.Rows.Add(item.GetAttribute("Name"), "");
            }

            if (!string.IsNullOrEmpty(cd["DailyLifeRecommend"]))
            {
                XmlElement dailyLifeRecommend = K12.Data.XmlHelper.LoadXml(cd["DailyLifeRecommend"]);
                tabControl1.Tabs[4].Text = dailyLifeRecommend.GetAttribute("Name");
            }

            if (_SR.Class == null)
            {
                this.Text = string.Format("日常生活表現評量   (  {0} 號  {1}  )", _SR.SeatNo, _SR.Name);
            }
            else
            {
                this.Text = string.Format("日常生活表現評量   (  {0} 班 {1} 號  {2}  )", _SR.Class.Name, _SR.SeatNo, _SR.Name);
            }

            //日常生活表現具體建議使用
            ReflashMorality();
            //努力程度
            ReflashEffortList();
            //日常行為表現
            ReflashDic();

            //權限
            btnSave.Visible = _permission.Editable;
            dgvDailyLifeRecommend.ReadOnly = !_permission.Editable;
            dgvDailyBehavior.ReadOnly = !_permission.Editable;
            dgvGroupActivity.ReadOnly = !_permission.Editable;
            dgvPublicService.ReadOnly = !_permission.Editable;
            dgvSchoolSpecial.ReadOnly = !_permission.Editable;

            List<int> colsDgvGroupActivity = new List<int>() { 1 };
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgvGroupActivity, colsDgvGroupActivity);

            List<int> colsDgvDailyBehavior = new List<int>() { 2 };
            Campus.Windows.DataGridViewImeDecorator dec2 = new Campus.Windows.DataGridViewImeDecorator(this.dgvDailyBehavior, colsDgvDailyBehavior);

        }

        private void BindData()
        {
            #region 更新資料
            if (_editorRecord == null || _editorRecord.TextScore == null || _editorRecord.TextScore.InnerXml == "") return;

            //DailyBehavior
            //<DailyBehavior Name="日常行為表現">
            //    <Item Name="愛整潔" Index="....." Degree="3"/>
            //    <Item Name="守秩序" Index="....." Degree="3"/>
            //</DailyBehavior>
            XmlElement node1 = (XmlElement)_editorRecord.TextScore.SelectSingleNode("DailyBehavior");

            if (node1 != null)
            {
                foreach (XmlElement item in node1.SelectNodes("Item"))
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
            XmlElement node2 = (XmlElement)_editorRecord.TextScore.SelectSingleNode("GroupActivity");

            if (node2 != null)
            {
                foreach (XmlElement item in node2.SelectNodes("Item"))
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
            XmlElement node3 = (XmlElement)_editorRecord.TextScore.SelectSingleNode("PublicService");

            if (node3 != null)
            {
                foreach (XmlElement item in node3.SelectNodes("Item"))
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
            XmlElement node4 = (XmlElement)_editorRecord.TextScore.SelectSingleNode("SchoolSpecial");

            if (node4 != null)
            {
                foreach (XmlElement item in node4.SelectNodes("Item"))
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
            XmlElement node5 = (XmlElement)_editorRecord.TextScore.SelectSingleNode("DailyLifeRecommend");

            if (node5 != null)
                dgvDailyLifeRecommend.Rows[0].Cells[0].Value = node5.GetAttribute("Description");
            #endregion
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (int.TryParse(cbSchoolYear.Text, out int year) && int.TryParse(cbSemester.Text, out int seme))
            {
                if (Mode == "NEW")
                {
                    List<JHMoralScoreRecord> MSRList = JHMoralScore.SelectByStudentIDs(new string[] { _SR.ID });

                    foreach (JHMoralScoreRecord each in MSRList)
                    {
                        if (each.SchoolYear.ToString() == cbSchoolYear.Text && each.Semester.ToString() == cbSemester.Text)
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("該學年度學期資料已經存在,無法新增");
                            return;
                        }
                    }

                    newSaveData();
                }
                else if (Mode == "UPDATA")
                {
                    updataSaveData();
                }
                else
                {
                    return;
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入正確的學年度/學期");
            }
        }

        private void newSaveData()
        {
            _editorRecord.RefStudentID = _SR.ID;
            _editorRecord.SchoolYear = int.Parse(cbSchoolYear.Text);
            _editorRecord.Semester = int.Parse(cbSemester.Text);

            sb_log.AppendLine("新增日常生活表現資料：");
            sb_log.AppendLine(string.Format("學生「{0}」學年度「{1}」學期「{2}」", _SR.Name, cbSchoolYear.Text, cbSemester.Text));
            sb_log.AppendLine("");

            SaveData(); //逗出XML

            try
            {
                string xyz = JHMoralScore.Insert(_editorRecord);

                List<JHMoralScoreRecord> list = JHMoralScore.SelectByStudentIDs(new string[] { _editorRecord.RefStudentID });
                foreach (JHMoralScoreRecord each in list)
                {
                    if (each.SchoolYear.ToString() == _editorRecord.SchoolYear.ToString() && each.Semester.ToString() == _editorRecord.Semester.ToString())
                    {
                        _editorRecord = each;
                    }
                }

            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("新增資料發生錯誤");
                return;
            }

            //新增後轉變為更新模式
            ApplicationLog.Log("日常生活表現", "新增", "student", _editorRecord.Student.ID, sb_log.ToString());

            FISCA.Presentation.Controls.MsgBox.Show("新增資料成功");
            cbSchoolYear.Enabled = false;
            cbSemester.Enabled = false;
            Mode = "UPDATA";
            CheckChange = false;
        }

        private void updataSaveData()
        {
            sb_log.AppendLine("更新日常生活表現資料：");
            sb_log.AppendLine(string.Format("學生「{0}」學年度「{1}」學期「{2}」", _SR.Name, cbSchoolYear.Text, cbSemester.Text));
            sb_log.AppendLine("");

            SaveData(); //逗出XML

            try
            {
                JHMoralScore.Update(_editorRecord);
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("更新資料發生錯誤");
                return;
            }

            ApplicationLog.Log("日常生活表現", "更新", "student", _editorRecord.Student.ID, sb_log.ToString());

            FISCA.Presentation.Controls.MsgBox.Show("更新資料成功");
            cbSchoolYear.Enabled = false;
            cbSemester.Enabled = false;
            CheckChange = false;
        }

        private void SaveData()
        {
            #region 更新儲存
            DSXmlHelper helper = new DSXmlHelper("TextScore");

            //DailyBehavior
            //<DailyBehavior Name="日常行為表現">
            //    <Item Name="愛整潔" Index="....." Degree="3"/>
            //    <Item Name="守秩序" Index="....." Degree="3"/>
            //</DailyBehavior>
            helper.AddElement("DailyBehavior").SetAttribute("Name", tabControl1.Tabs[0].Text);
            sb_log.AppendLine(tabControl1.Tabs[0].Text + "：");
            foreach (DataGridViewRow row in dgvDailyBehavior.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("DailyBehavior", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Index", "" + row.Cells[1].Value);
                node.SetAttribute("Degree", "" + row.Cells[2].Value);

                sb_log.AppendLine(string.Format("項目「{0}」表現程度「{1}」", "" + row.Cells[0].Value, "" + row.Cells[2].Value));
            }

            sb_log.AppendLine("");

            //GroupActivity
            //<GroupActivity Name="團體活動表現">
            //    <Item Name="社團活動" Degree="1" Description=".....">
            //    <Item Name="學校活動" Degree="2" Description=".....">
            //</GroupActivity>
            helper.AddElement("GroupActivity").SetAttribute("Name", tabControl1.Tabs[1].Text);
            sb_log.AppendLine(tabControl1.Tabs[1].Text + "：");
            foreach (DataGridViewRow row in dgvGroupActivity.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("GroupActivity", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Degree", "" + row.Cells[1].Value);
                node.SetAttribute("Description", "" + row.Cells[2].Value);

                sb_log.AppendLine(string.Format("項目「{0}」努力程度「{1}」文字描述「{2}」", "" + row.Cells[0].Value, "" + row.Cells[1].Value, "" + row.Cells[2].Value));
            }

            sb_log.AppendLine("");

            //PublicService
            //<PublicService Name="公共服務表現">
            //    <Item Name="校內服務" Description=".....">
            //    <Item Name="社區服務" Description=".....">
            //</PublicService>
            helper.AddElement("PublicService").SetAttribute("Name", tabControl1.Tabs[2].Text);

            sb_log.AppendLine(tabControl1.Tabs[2].Text + "：");
            foreach (DataGridViewRow row in dgvPublicService.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("PublicService", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Description", "" + row.Cells[1].Value);

                sb_log.AppendLine(string.Format("項目「{0}」文字描述「{1}」", "" + row.Cells[0].Value, "" + row.Cells[1].Value));
            }

            sb_log.AppendLine("");

            //SchoolSpecial
            //<SchoolSpecial Name="校內外時特殊表現">
            //    <Item Name="校外特殊表現" Description=".....">
            //    <Item Name="校內特殊表現" Description=".....">
            //</SchoolSpecial>
            helper.AddElement("SchoolSpecial").SetAttribute("Name", tabControl1.Tabs[3].Text);
            sb_log.AppendLine(tabControl1.Tabs[3].Text + "：");

            foreach (DataGridViewRow row in dgvSchoolSpecial.Rows)
            {
                if (row.IsNewRow) continue;

                XmlElement node = helper.AddElement("SchoolSpecial", "Item");
                node.SetAttribute("Name", "" + row.Cells[0].Value);
                node.SetAttribute("Description", "" + row.Cells[1].Value);

                sb_log.AppendLine(string.Format("項目「{0}」文字描述「{1}」", "" + row.Cells[0].Value, "" + row.Cells[1].Value));
            }

            sb_log.AppendLine("");

            //DailyLifeRecommend
            //<DailyLifeRecommend Name="日常生活表現具體建議" Description=".....">
            XmlElement anode = helper.AddElement("DailyLifeRecommend");
            anode.SetAttribute("Name", tabControl1.Tabs[4].Text);
            anode.SetAttribute("Description", "" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value);

            sb_log.AppendLine(string.Format("{0}：\n建議內容「{1}」", tabControl1.Tabs[4].Text, "" + dgvDailyLifeRecommend.Rows[0].Cells[0].Value));


            _editorRecord.TextScore = helper.BaseElement;

            #endregion
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            if (CheckChange)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("資料已經修改是否關閉?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
                if (dr == DialogResult.Yes)
                {
                    this.Close();
                }
                else
                {
                    return;
                }
            }
            else
            {
                this.Close();
            }
        }

        /// <summary>
        /// 檢查儲存按鈕是否可以按。當格子裡沒有錯誤的值時才Enabled。
        /// </summary>       
        private void CheckSaveButtonEnabled()
        {
            this.btnSave.Enabled = !this.inputErrors.ContainsValue(false);
        }

        //<DailyBehavior Name="日常行為表現">
        //    <Item Name="愛整潔" Index="....."/>
        //    <PerformanceDegree>
        //        <Mapping Degree="4" Desc="完全符合"/>
        //        <Mapping Degree="3" Desc="大部份符合"/>
        //        <Mapping Degree="2" Desc="部份符合"/>
        //    </PerformanceDegree>
        //</DailyBehavior>

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

        private void ReflashEffortList()
        {
            #region 努力程度對照表
            EffortList.Clear();
            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["努力程度對照表"];
            if (!string.IsNullOrEmpty(cd["xml"]))
            {
                XmlElement element = K12.Data.XmlHelper.LoadXml(cd["xml"]);

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
            XmlElement node = K12.Data.XmlHelper.LoadXml(cd["DailyBehavior"]);
            foreach (XmlElement item in node.SelectNodes("PerformanceDegree/Mapping"))
            {
                if (!dic.ContainsKey(item.GetAttribute("Degree")))
                {
                    dic.Add(item.GetAttribute("Degree"), item.GetAttribute("Desc"));
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("表現程度代碼表,代碼重覆");
                }
            }
            #endregion
        }

        private void dgvGroupActivity_CurrentCellDirtyStateChanged_1(object sender, EventArgs e)
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

        private void dgvDailyBehavior_CurrentCellDirtyStateChanged_1(object sender, EventArgs e)
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

        private void dgvDailyLifeRecommend_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgvDailyLifeRecommend.EndEdit();
            dgvDailyLifeRecommend.BeginEdit(false);
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

        private void cbSchoolYear_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(cbSchoolYear.Text, out int year) && int.TryParse(cbSemester.Text, out int seme))
            {
                JHMoralScoreRecord CheckMSR = JHMoralScore.SelectBySchoolYearAndSemester(_PrimaryKey, year, seme);
                DataListener.SuspendListen(); //終止變更判斷
                if (CheckMSR == null)
                {
                    Mode = "NEW";
                    _editorRecord = new JHMoralScoreRecord();
                    SyndLoad();
                }
                else
                {
                    Mode = "UPDATA";
                    _editorRecord = CheckMSR;

                    SyndLoad();

                    BindData();
                }
                DataListener.Reset();
                DataListener.ResumeListen();
                inputErrors.Clear();
                btnSave.Enabled = true;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入正確的學年度/學期");
            }
        }

        private void cbSemester_TextChanged(object sender, EventArgs e)
        {
            if (int.TryParse(cbSchoolYear.Text, out int year) && int.TryParse(cbSemester.Text, out int seme))
            {
                JHMoralScoreRecord CheckMSR = JHMoralScore.SelectBySchoolYearAndSemester(_PrimaryKey, year, seme);
                DataListener.SuspendListen(); //終止變更判斷
                if (CheckMSR == null)
                {
                    Mode = "NEW";
                    _editorRecord = new JHMoralScoreRecord();
                    SyndLoad();
                }
                else
                {
                    Mode = "UPDATA";
                    _editorRecord = CheckMSR;

                    SyndLoad();

                    BindData();
                }
                DataListener.Reset();
                DataListener.ResumeListen();
                inputErrors.Clear();
                btnSave.Enabled = true;
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("請輸入正確的學年度/學期");
            }
        }
    }
}
