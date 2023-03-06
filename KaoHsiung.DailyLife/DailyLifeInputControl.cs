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
using FISCA.LogAgent;
using JHSchool;

namespace KaoHsiung.DailyLife
{
    //權限代碼：JHSchool.StuAdmin.Ribbon0025

    //<DailyLifeInputConfig SchoolYear="97" Semester="2">
    //    <InputTimeControl>
    //        <Time End="2009/5/10 15:00" Grade="7" Start="2009/5/1 15:00" />
    //        <Time End="2009/5/10 15:00" Grade="8" Start="2009/5/1 15:00" />
    //        <Time End="2009/5/10 15:00" Grade="9" Start="2009/5/1 15:00" />
    //    </InputTimeControl>
    //</DailyLifeInputConfig>

    public partial class DailyLifeInputControl : FISCA.Presentation.Controls.BaseForm
    {
        private const string DateTimeFormat = "yyyy/MM/dd HH:mm";

        StringBuilder sb1 = new StringBuilder();

        /// <summary>
        /// 日常生活表現輸入時間控制的組態名稱。 
        /// </summary>
        public const string ConfigName = "DailyLifeInputConfig";

        public DailyLifeInputControl()
        {
            InitializeComponent();
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgvTimes);
        }

        private void DailyLifeInputControl_Load(object sender, EventArgs e)
        {
            lblSemester.Text = string.Format("{0}學年度　第{1}學期", School.DefaultSchoolYear, School.DefaultSemester);

            List<string> cols = new List<string>() { "年級" , "開始時間" , "結束時間" };
            Campus.Windows.DataGridViewImeDecorator dec = new Campus.Windows.DataGridViewImeDecorator(this.dgvTimes, cols);

            //先將 Grid 填入此學校有的年級。
            FillGridViewGradeYear();

            //將對應年級的時間填入。
            FillTimes();
        }

        private void FillTimes()
        {
            School.Configuration.Sync(ConfigName);
            ConfigData cd = School.Configuration[ConfigName];

            XmlElement config = cd.GetXml("XmlData", null);

            //沒有資料就不顯示資料。
            if (config == null) return;

            string schoolyear = config.GetAttribute("SchoolYear");
            string semester = config.GetAttribute("Semester");
            sb1.AppendLine("已對「日常生活表現輸入時間」進行修改。");
            sb1.AppendLine("修改前資料：");
            sb1.Append("學年度「" + schoolyear + "」");
            sb1.AppendLine("學期「" + semester + "」");
            foreach (XmlElement each in config.SelectNodes("InputTimeControl/Time"))
            {
                XmlHelper eachTime = new XmlHelper(each);
                string grade = eachTime.GetString("@Grade");
                string startTime = eachTime.GetString("@Start");
                string endTime = eachTime.GetString("@End");

                foreach (DataGridViewRow eachRow in dgvTimes.Rows)
                {
                    string rowgrade = eachRow.Cells[chGradeYear.Index].Value + "";

                    if (rowgrade == grade)
                    {
                        eachRow.Cells[chStartTime.Index].Value = startTime;
                        eachRow.Cells[chEndTime.Index].Value = endTime;

                        sb1.Append("年級「" + rowgrade + "」");
                        sb1.Append("開始輸入時間「" + startTime + "」");
                        sb1.AppendLine("結束輸入時間「" + endTime + "」");
                    }
                }
            }
        }

        private void FillGridViewGradeYear()
        {
            foreach (int each in GroupGradeYear())
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dgvTimes, each.ToString(), "", "");
                dgvTimes.Rows.Add(row);
            }
        }

        private List<int> GroupGradeYear()
        {
            Dictionary<int, string> years = new Dictionary<int, string>();
            foreach (ClassRecord each in Class.Instance.Items)
            {
                int year;

                if (!int.TryParse(each.GradeYear, out year))
                    continue;

                if (!years.ContainsKey(year))
                    years.Add(year, each.GradeYear);
            }

            List<int> intyears = new List<int>(years.Keys);
            intyears.Sort();

            return intyears;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (IsDataValidity())
            {
                #region 資料正確,進行儲存

                StringBuilder sb2 = new StringBuilder();

                XmlCreator dailylife = new XmlCreator();
                dailylife.CreateStartElement("DailyLifeInputConfig");
                {
                    sb1.AppendLine("修改後資料：");
                    sb2.Append("學年度「" + School.DefaultSchoolYear + "」");
                    sb2.AppendLine("學期「" + School.DefaultSemester + "」");

                    dailylife.CreateAttribute("SchoolYear", School.DefaultSchoolYear);
                    dailylife.CreateAttribute("Semester", School.DefaultSemester);

                    dailylife.CreateStartElement("InputTimeControl");
                    {
                        foreach (DataGridViewRow each in dgvTimes.Rows)
                        {
                            dailylife.CreateStartElement("Time");
                            dailylife.CreateAttribute("Grade", each.Cells[chGradeYear.Index].Value + "");
                            dailylife.CreateAttribute("Start", each.Cells[chStartTime.Index].Value + "");
                            dailylife.CreateAttribute("End", each.Cells[chEndTime.Index].Value + "");
                            dailylife.CreateEndElement();

                            sb2.Append("年級「" + each.Cells[chGradeYear.Index].Value + "」");
                            sb2.Append("開始輸入時間「" + each.Cells[chStartTime.Index].Value + "」");
                            sb2.AppendLine("結束輸入時間「" + each.Cells[chEndTime.Index].Value + "」");

                        }
                    }
                    dailylife.CreateEndElement();
                }
                dailylife.CreateEndElement();

                ConfigData cd = School.Configuration[ConfigName];
                cd.SetXml("XmlData", dailylife.GetAsXmlElement());
                cd.Save();

                ApplicationLog.Log("日常生活表現輸入時間", "修改", sb1.ToString() + sb2.ToString());
                MsgBox.Show("儲存成功!!");
                DialogResult = DialogResult.OK;

                #endregion
            }
            else
            {
                MsgBox.Show("畫面中含有不正確資料。");
                DialogResult = DialogResult.None;
            }
        }

        private bool IsDataValidity()
        {
            bool valid = true;
            foreach (DataGridViewRow each in dgvTimes.Rows)
            {
                if (!string.IsNullOrEmpty(each.ErrorText))
                {
                    valid = false;
                }

                foreach (DataGridViewCell eachCell in each.Cells)
                {
                    if (!string.IsNullOrEmpty(eachCell.ErrorText))
                    {
                        valid = false;
                    }
                }

                if (!valid) break;
            }

            return valid;
        }

        private void dgvTimes_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //dgvTimes.BeginEdit(true);
        }

        /// <summary>
        /// 開始&結束日期是否有錯誤
        /// </summary>
        private void dgvTimes_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            DataGridViewRow row = dgvTimes.Rows[e.RowIndex];

            string startTime = row.Cells[chStartTime.Index].Value + "";
            string endTime = row.Cells[chEndTime.Index].Value + "";

            row.ErrorText = "";
            if (string.IsNullOrEmpty(startTime) && string.IsNullOrEmpty(endTime))
            {
                //這裡沒有程式。
            }
            else if (!string.IsNullOrEmpty(startTime) && !string.IsNullOrEmpty(endTime))
            {
                DateTime? objStart = DateTimeHelper.Parse(startTime);
                DateTime? objEnd = DateTimeHelper.Parse(endTime);

                if (objStart.HasValue && objEnd.HasValue)
                {
                    if (objStart.Value >= objEnd.Value)
                        row.ErrorText = "截止時間必須在開始時間之後。";
                }
            }
            else
                row.ErrorText = "請輸入正確的時間限制資料(必需同時有資料或同時沒有資料)。";
        }

        private void dgvTimes_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //ValidateCellData(e.ColumnIndex, e.RowIndex, e.FormattedValue + "");
        }

        private void dgvTimes_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ValidateCellData(e.ColumnIndex, e.RowIndex, dgvTimes.Rows[e.RowIndex].Cells[e.ColumnIndex].Value + "");
            //TryToCorrectData(e.ColumnIndex, e.RowIndex);
        }

        private void ValidateCellData(int columnIndex, int rowIndex, string value)
        {
            if (columnIndex == chStartTime.Index || columnIndex == chEndTime.Index)
            {
                DataGridViewCell cell = dgvTimes.Rows[rowIndex].Cells[columnIndex];
                cell.ErrorText = "";
                if (string.IsNullOrEmpty(value)) //沒有資料就不作任何檢查。
                    return;

                DateTime dt;
                if (!DateTime.TryParse(value, out dt))
                {
                    cell.ErrorText = "日期格式錯誤。";
                }
                else
                {
                    cell.Value = dt.ToString(DateTimeFormat);
                }
            }
        }

        private void dgvTimes_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            //TryToCorrectData(e.ColumnIndex, e.RowIndex);
        }

        private void TryToCorrectData(int columnIndex, int rowIndex)
        {
            if (columnIndex == chStartTime.Index || columnIndex == chEndTime.Index)
            {
                DataGridViewRow row = dgvTimes.Rows[rowIndex];
                row.Cells[columnIndex].ErrorText = string.Empty;
                string time = row.Cells[columnIndex].Value + "";

                if (string.IsNullOrEmpty(time)) //沒有資料就不作任何檢查。
                    return;

                DateTime? objStart = DateTimeHelper.ParseGregorian(time, PaddingMethod.First);

                if (objStart.HasValue)
                    row.Cells[columnIndex].Value = objStart.Value.ToString(DateTimeFormat);
            }
        }
    }
}
