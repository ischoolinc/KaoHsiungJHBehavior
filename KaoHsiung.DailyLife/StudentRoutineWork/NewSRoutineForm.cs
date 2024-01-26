using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Presentation;
using Aspose.Words;
using K12.Data;
using System.IO;
using System.Diagnostics;
using JHSchool.Data;
using System.Xml;
using SmartSchool.ePaper;
using K12EmergencyContact.DAO;





namespace KaoHsiung.DailyLife.StudentRoutineWork
{
    public partial class NewSRoutineForm : BaseForm
    {

        private BackgroundWorker BGW = new BackgroundWorker();
        //主文件
        private Document _doc;
        //單頁範本
        private Document _template;
        //移動使用
        private Run _run;

        List<string> DLBList1 = new List<string>();
        List<string> DLBList2 = new List<string>();

        //取得StudentIDList

        List<string> StudentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;

        //DailyBehavior
        //DailyLifeRecommend
        //GroupActivity
        //PublicService
        //SchoolSpecial
        Dictionary<string, string> TieDic1 = new Dictionary<string, string>(); Dictionary<string, int> DicSummaryIndex = new Dictionary<string, int>();
        Dictionary<string, string> UpdateCoddic = new Dictionary<string, string>();



        List<SemesterSLR> SLRNameList { get; set; }

        bool PrintSaveFile = false;
        bool PrintUpdateStudentFile = false;

        /// <summary>
        /// 學生電子報表
        /// </summary>
        SmartSchool.ePaper.ElectronicPaper paperForStudent { get; set; }

        Dictionary<StudentDataObj, Document> StudentSaveDic { get; set; }

        public NewSRoutineForm()
        {
            InitializeComponent();
        }

        private void NewSRoutineForm_Load(object sender, EventArgs e)
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {

            MotherForm.SetStatusBarMessage("開始列印學生訓導記錄表...");
            btnSave.Enabled = false;

            _doc = new Document();
            _doc.Sections.Clear(); //清空此Document

            PrintSaveFile = cbPrintSaveFile.Checked;
            PrintUpdateStudentFile = cbPrintUpdateStudentFile.Checked;

            _template = new Document(new MemoryStream(KaoHsiung.DailyLife.Properties.Resources.記錄表Word));

            SetNameIndex();

            BGW.RunWorkerAsync();
        }

        /// <summary>
        /// 背景模式
        /// </summary>
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            paperForStudent = new SmartSchool.ePaper.ElectronicPaper(School.DefaultSchoolYear + "學生訓導記錄表", School.DefaultSchoolYear, School.DefaultSemester, SmartSchool.ePaper.ViewerType.Student);
            StudentSaveDic = new Dictionary<StudentDataObj, Document>();

            Dictionary<string, udt_K12EmergencyContact> emergencyContact = GetEmergencyContactor(StudentIDList); // 緊急聯絡人

            StudentInfo Data = new StudentInfo();

            foreach (string student in Data.DicStudent.Keys)
            {
                #region 學生

                //取得學生資料物件
                StudentDataObj obj = Data.DicStudent[student];
                //取得範本樣式
                Document PageOne = (Document)_template.Clone(true);
                //???
                _run = new Run(PageOne);
                //可建構的...
                DocumentBuilder builder = new DocumentBuilder(PageOne);
                DocumentBuilder builderX = new DocumentBuilder(_template);

                #region 資料MailMerge第一步

                List<string> name = new List<string>();
                List<string> value = new List<string>();

                //  2016/3/30   穎驊處理完畢
                name.Add("緊急");
                if (emergencyContact.ContainsKey(obj.StudentRecord.ID))
                {
                    value.Add(emergencyContact[obj.StudentRecord.ID].ContactName);
                }
                else
                {
                    value.Add("");

                }


                name.Add("學校名稱");
                value.Add(School.ChineseName);

                name.Add("學號");
                value.Add(obj.StudentRecord.StudentNumber);

                name.Add("姓名");
                value.Add(obj.StudentRecord.Name);

                name.Add("性別");
                value.Add(obj.StudentRecord.Gender);

                name.Add("身分證");
                value.Add(obj.StudentRecord.IDNumber);

                name.Add("生日");
                value.Add(obj.StudentRecord.Birthday.HasValue ? obj.StudentRecord.Birthday.Value.ToShortDateString() : "");

                name.Add("出生");
                value.Add(obj.StudentRecord.BirthPlace);

                name.Add("監護");
                value.Add(obj.CustodianName);

                name.Add("戶籍");
                value.Add(obj.AddressPermanent);

                name.Add("電話1");
                value.Add(obj.PhonePermanent);




                name.Add("通訊");
                value.Add(obj.AddressMailing);

                name.Add("電話2");
                value.Add(obj.PhoneContact);

                name.Add("畢業國小");
                value.Add(obj.UpdataGraduateSchool);

                name.Add("入學日期");
                value.Add(obj.UpdataADDate);

                name.Add("入學文號");
                value.Add(obj.UpdataADNumber);

                name.Add("一上");
                value.Add(obj.GradeYear11);

                name.Add("一下");
                value.Add(obj.GradeYear12);

                name.Add("二上");
                value.Add(obj.GradeYear21);

                name.Add("二下");
                value.Add(obj.GradeYear22);

                name.Add("三上");
                value.Add(obj.GradeYear31);

                name.Add("三下");
                value.Add(obj.GradeYear32);

                name.Add("日期");
                value.Add(DateTime.Now.ToString("yyyy/MM/dd"));

                name.Add("時間");
                value.Add(DateTime.Now.ToString("HH:mm:ss"));

                List<string> value2 = new List<string>();
                foreach (string each in value)
                {
                    if (!string.IsNullOrEmpty(each))
                        value2.Add(SurrogatePairString(each));
                    else
                        value2.Add(each);
                }
                PageOne.MailMerge.Execute(name.ToArray(), value2.ToArray());
                #endregion

                #region 異動處理

                //移動到(MergeField)
                builder.MoveToMergeField("異動");
                //取得目前Cell
                Cell UpdateRecordCell = (Cell)builder.CurrentParagraph.ParentNode;
                //取得目前Row
                Row row = (Row)builder.CurrentParagraph.ParentNode.ParentNode;

                //建立新行(依異動筆數)
                for (int x = 1; x < obj.ListUpdateRecord.Count; x++)
                {
                    (UpdateRecordCell.ParentNode.ParentNode as Table).InsertAfter(row.Clone(true), UpdateRecordCell.ParentNode);
                }

                foreach (JHUpdateRecordRecord updateRecord in obj.ListUpdateRecord)
                {
                    List<string> list = new List<string>();
                    list.Add(updateRecord.SchoolYear.HasValue ? updateRecord.SchoolYear.Value.ToString() : "");
                    list.Add(updateRecord.Semester.HasValue ? updateRecord.Semester.Value.ToString() : "");
                    list.Add(updateRecord.UpdateDate);
                    list.Add(updateRecord.ADDate);
                    list.Add(GetUpdateRecordCode(updateRecord.UpdateCode));
                    list.Add(updateRecord.ADNumber);
                    list.Add(updateRecord.Comment);

                    foreach (string UpdateName in list)
                    {
                        //寫入
                        Write(UpdateRecordCell, UpdateName);
                        if (UpdateRecordCell.NextSibling != null) //是否最後一格
                            UpdateRecordCell = UpdateRecordCell.NextSibling as Cell; //下一格
                    }
                    Row Nextrow = UpdateRecordCell.ParentRow.NextSibling as Row; //取得下一個Row
                    UpdateRecordCell = Nextrow.FirstCell; //第一格Cell           
                }



                #endregion

                #region 日常生活處理

                GetBehaviorConfig(); //取得設定

                PageOne.MailMerge.Execute(TieDic1.Keys.ToArray(), TieDic1.Values.ToArray());


                //移動到(MergeField)
                builder.MoveToMergeField("設定1");
                Cell setupCell = (Cell)builder.CurrentParagraph.ParentNode;
                foreach (string each in DLBList1)
                {
                    Write(setupCell, each); //填入愛整潔
                    if (setupCell.NextSibling != null)
                    {
                        setupCell = setupCell.NextSibling as Cell; //取得下一格
                    }
                }

                int RowNull = 0;

                builder.MoveToMergeField("日1");
                Cell MoralScore1 = (Cell)builder.CurrentParagraph.ParentNode;

                foreach (string moralScore in obj.TextScoreDic.Keys)
                {
                    Write(MoralScore1, moralScore); //填入學年度
                    MoralScore1 = MoralScore1.NextSibling as Cell; //取得下一格

                    foreach (string BehaviorConfigName1 in DLBList1)
                    {
                        if (obj.TextScoreDic[moralScore].ContainsKey(BehaviorConfigName1)) //如果包含資料
                        {
                            Write(MoralScore1, obj.TextScoreDic[moralScore][BehaviorConfigName1]);
                            if (MoralScore1.NextSibling != null)
                            {
                                MoralScore1 = MoralScore1.NextSibling as Cell;
                            }
                        }
                    }

                    Row Nextrow = MoralScore1.ParentRow.NextSibling as Row; //取得下一個Row
                    MoralScore1 = Nextrow.FirstCell; //第一格Cell

                    RowNull++;
                    if (RowNull >= 6)
                        break;
                }

                builder.MoveToMergeField("日2");
                RowNull = 0;
                Cell MoralScore2 = (Cell)builder.CurrentParagraph.ParentNode;
                foreach (string moralScore in obj.TextScoreDic.Keys)
                {
                    Write(MoralScore2, moralScore); //學年度
                    MoralScore2 = MoralScore2.NextSibling as Cell;

                    foreach (string BehaviorConfigName2 in DLBList2)
                    {
                        if (obj.TextScoreDic[moralScore].ContainsKey(BehaviorConfigName2))
                        {
                            if (obj.TextScoreDic[moralScore][BehaviorConfigName2] != "GroupActivity")
                            {
                                Write(MoralScore2, obj.TextScoreDic[moralScore][BehaviorConfigName2]);
                            }
                        }

                        //如果下一隔不是空的
                        if (MoralScore2.NextSibling != null)
                        {
                            MoralScore2 = MoralScore2.NextSibling as Cell;
                        }
                    }

                    Row Nextrow = MoralScore2.ParentRow.NextSibling as Row; //取得下一個Row
                    MoralScore2 = Nextrow.FirstCell; //第一格Cell

                    RowNull++;
                    if (RowNull >= 6)
                        break;

                }
                #endregion

                #region 缺曠統計處理

                builder.MoveToMergeField("統1");
                RowNull = 0;
                Cell MoralScore3 = (Cell)builder.CurrentParagraph.ParentNode;

                foreach (string moralScore in obj.SummaryDic.Keys)
                {
                    //填入學期
                    Write(MoralScore3, moralScore);

                    if (obj.SchoolDay.ContainsKey(moralScore))
                    {
                        Cell MoralScore5 = MoralScore3.NextSibling as Cell;
                        Write(MoralScore5, obj.SchoolDay[moralScore]); //寫入上課天數
                    }

                    foreach (string SummaryName in obj.SummaryDic[moralScore].Keys)
                    {
                        int index = GetSummaryIndex(SummaryName);

                        //如果是0,就是沒有值
                        if (index == 0)
                            continue;
                        //取得MoralScore3為基準的 index 格
                        Cell MoralScore4 = GetMoveRightCell(MoralScore3, index);
                        //填入值
                        if (obj.SummaryDic[moralScore][SummaryName] != "0")
                        {
                            Write(MoralScore4, obj.SummaryDic[moralScore][SummaryName]);
                        }
                    }

                    Row Nextrow2 = MoralScore3.ParentRow.NextSibling as Row; //取得下一個Row
                    MoralScore3 = Nextrow2.FirstCell; //第一格Cell

                    RowNull++;
                    if (RowNull >= 6)
                        break;
                }

                #endregion

                #region 獎懲明細處理

                builder.MoveToMergeField("獎懲");
                Cell MeritDemeritCell = (Cell)builder.CurrentParagraph.ParentNode;
                //取得目前Row
                Row Derow = (Row)builder.CurrentParagraph.ParentNode.ParentNode;

                int MeritDemeritIndex = obj.ListMerit.Count;

                foreach (DemeritRecord demerit in obj.ListDeMerit)
                {
                    if (demerit.Cleared != "是")
                        MeritDemeritIndex++;
                }

                //建立新行(依異動筆數)
                for (int x = 1; x < MeritDemeritIndex; x++)
                {
                    (MeritDemeritCell.ParentNode.ParentNode as Table).InsertAfter(Derow.Clone(true), MeritDemeritCell.ParentNode);
                }

                foreach (MeritRecord merit in obj.ListMerit)
                {
                    #region 獎勵
                    string MeritSchoolYearSemerit = merit.SchoolYear.ToString() + "/" + merit.Semester.ToString();
                    string day = merit.OccurDate.ToShortDateString();
                    string A = merit.MeritA.HasValue ? merit.MeritA.Value.ToString() : "";
                    string B = merit.MeritB.HasValue ? merit.MeritB.Value.ToString() : "";
                    string C = merit.MeritC.HasValue ? merit.MeritC.Value.ToString() : "";
                    string Reason = merit.Reason;
                    string remark = merit.Remark;

                    Cell MeritCellDay = GetMoveRightCell(MeritDemeritCell, 1);
                    Cell MeritCellA = GetMoveRightCell(MeritDemeritCell, 2);
                    Cell MeritCellB = GetMoveRightCell(MeritDemeritCell, 3);
                    Cell MeritCellC = GetMoveRightCell(MeritDemeritCell, 4);
                    Cell MeritCellReason = GetMoveRightCell(MeritDemeritCell, 8);
                    Cell MeritCellRemark = GetMoveRightCell(MeritDemeritCell, 9);

                    Write(MeritDemeritCell, MeritSchoolYearSemerit); //學年度
                    Write(MeritCellDay, day);
                    if (A != "0")
                    {
                        Write(MeritCellA, A);
                    }
                    if (B != "0")
                    {
                        Write(MeritCellB, B);
                    }
                    if (C != "0")
                    {
                        Write(MeritCellC, C);
                    }

                    Write(MeritCellReason, Reason);
                    Write(MeritCellRemark, remark);

                    Row Nextrow2 = MeritDemeritCell.ParentRow.NextSibling as Row; //取得下一個Row
                    if (Nextrow2 != null)
                    {
                        MeritDemeritCell = Nextrow2.FirstCell; //第一格Cell
                    }
                    #endregion
                }

                foreach (DemeritRecord demerit in obj.ListDeMerit)
                {
                    if (demerit.Cleared == "是")
                        continue;

                    #region 懲戒
                    string DemeritSchoolYearSemerit = demerit.SchoolYear.ToString() + "/" + demerit.Semester.ToString();
                    string day = demerit.OccurDate.ToShortDateString();
                    string A = demerit.DemeritA.HasValue ? demerit.DemeritA.Value.ToString() : "";
                    string B = demerit.DemeritB.HasValue ? demerit.DemeritB.Value.ToString() : "";
                    string C = demerit.DemeritC.HasValue ? demerit.DemeritC.Value.ToString() : "";
                    string Reason = demerit.Reason;
                    string remark = demerit.Remark;

                    Cell DemeritCellDay = GetMoveRightCell(MeritDemeritCell, 1);
                    Cell DemeritCellA = GetMoveRightCell(MeritDemeritCell, 5);
                    Cell DemeritCellB = GetMoveRightCell(MeritDemeritCell, 6);
                    Cell DemeritCellC = GetMoveRightCell(MeritDemeritCell, 7);
                    Cell DemeritCellReason = GetMoveRightCell(MeritDemeritCell, 8);
                    Cell DemeritCellRemark = GetMoveRightCell(MeritDemeritCell, 9);

                    Write(MeritDemeritCell, DemeritSchoolYearSemerit); //學年度
                    Write(DemeritCellDay, day);
                    if (A != "0")
                    {
                        Write(DemeritCellA, A);
                    }
                    if (B != "0")
                    {
                        Write(DemeritCellB, B);
                    }
                    if (C != "0")
                    {
                        Write(DemeritCellC, C);
                    }
                    Write(DemeritCellReason, Reason);
                    Write(DemeritCellRemark, remark);
                    Row Nextrow2 = MeritDemeritCell.ParentRow.NextSibling as Row; //取得下一個Row
                    if (Nextrow2 != null)
                    {
                        MeritDemeritCell = Nextrow2.FirstCell; //第一格Cell
                    }
                    else
                    {
                        foreach (Cell each in MeritDemeritCell.ParentRow.Cells)
                        {
                            each.CellFormat.Borders.Bottom.LineWidth = 1.5;
                        }
                    }
                    #endregion
                }

                #endregion

                builder.MoveToMergeField("學習");
                Cell sprder = (Cell)builder.CurrentParagraph.ParentNode;
                foreach (Cell ce_ll in sprder.ParentRow.Cells)
                {
                    ce_ll.CellFormat.Borders.Top.LineWidth = 1.5;
                }

                #region 服務學習時數

                Dictionary<string, decimal> SchoolSLRDic = new Dictionary<string, decimal>();
                Dictionary<string, SemesterSLR> SLRNameDic = new Dictionary<string, SemesterSLR>();
                SLRNameList = new List<SemesterSLR>();

                foreach (SLRecord slr in obj.ListSLR)
                {
                    string sString = string.Format("「{0}」學年度　第「{1}」學期", slr.SchoolYear.ToString(), slr.Semester.ToString());

                    SemesterSLR s = new SemesterSLR(slr);
                    if (!SLRNameDic.ContainsKey(sString))
                    {
                        SLRNameDic.Add(sString, s);
                        SLRNameList.Add(s);
                    }

                    if (SchoolSLRDic.ContainsKey(sString))
                    {
                        SchoolSLRDic[sString] += slr.Hours;
                    }
                    else
                    {
                        SchoolSLRDic.Add(sString, 0);
                        SchoolSLRDic[sString] += slr.Hours;
                    }
                }

                //排序學年度學期 5/15
                SLRNameList.Sort(SortSLR);

                builder.MoveToMergeField("服務");

                Cell SLRCell = (Cell)builder.CurrentParagraph.ParentNode;
                Row SLRrow = (Row)builder.CurrentParagraph.ParentNode.ParentNode;

                for (int x = 1; x < SchoolSLRDic.Count; x++)
                {
                    (SLRCell.ParentNode.ParentNode as Table).InsertAfter(SLRrow.Clone(true), SLRCell.ParentNode);
                }

                foreach (SemesterSLR each in SLRNameList)
                {
                    string s = string.Format("「{0}」學年度　第「{1}」學期", each.SchoolYear.ToString(), each.Semester.ToString());
                    //學年期
                    Write(SLRCell, s);

                    //時數
                    Cell SLRCellNext = GetMoveRightCell(SLRCell, 1);
                    Write(SLRCellNext, "共「" + SchoolSLRDic[s] + "」小時");

                    Row Nextrow2 = SLRCell.ParentRow.NextSibling as Row;
                    if (Nextrow2 != null)
                    {
                        SLRCell = Nextrow2.FirstCell; //第一格Cell
                    }
                    else
                    {
                        foreach (Cell ce_ll in SLRCell.ParentRow.Cells)
                        {
                            ce_ll.CellFormat.Borders.Bottom.LineWidth = 1.5;
                        }
                    }
                }

                #endregion

                MemoryStream stream = new MemoryStream();
                //PageOne.Save(stream, SaveFormat.Doc);
                //paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, student));

                PageOne.Save(stream, SaveFormat.Doc);

                if (cbIsPdf.Checked)
                {
                    MemoryStream stream_pdf = new MemoryStream();

                    stream_pdf = (MemoryStream)Aspose.IO.Tools.SavePDFtoStream(stream);

                    paperForStudent.Append(new PaperItem(PaperFormat.AdobePdf, stream_pdf, student));
                }

                if (cbIsDoc.Checked)
                {
                    paperForStudent.Append(new PaperItem(PaperFormat.Office2003Doc, stream, student));
                }


                StudentSaveDic.Add(obj, PageOne);

                //將PageOne加入主文件內
                _doc.Sections.Add(_doc.ImportNode(PageOne.FirstSection, true));

                #endregion
            }

            //如果有打勾則上傳電子報表
            if (PrintUpdateStudentFile)
                SmartSchool.ePaper.DispatcherProvider.Dispatch(paperForStudent);

            e.Result = _doc;
        }

        /// <summary>
        /// 將特殊字用空白表示
        /// </summary>
        public string SurrogatePairString(string input)
        {
            string value = "";
            int idx = 0;
            foreach (char c in input)
            {
                if (char.IsSurrogatePair(input, idx) || char.IsSurrogate(c))
                {
                    value += " ";
                }
                else
                {
                    value += c;
                }
                idx++;
            }
            return value;
        }



        public Dictionary<string, udt_K12EmergencyContact> GetEmergencyContactor(List<string> list)
        {
            #region 取得選擇學生之緊急連絡人

            Dictionary<string, udt_K12EmergencyContact> emergencyContactor = new Dictionary<string, udt_K12EmergencyContact>();

            foreach (udt_K12EmergencyContact each in UDTTransfer.GetStudentK12EmergencyContactByStudentIDList(list))
            {
                string studID = each.RefStudentID.ToString();

                if (!emergencyContactor.ContainsKey(studID))
                    emergencyContactor.Add(studID, each);
            }

            return emergencyContactor;

            #endregion
        }

        public int SortSLR(SemesterSLR s1, SemesterSLR s2)
        {
            string s1_t = s1.SchoolYear.ToString().PadLeft(3, '0');
            s1_t += s1.Semester.ToString().PadLeft(1, '0');
            string s2_t = s2.SchoolYear.ToString().PadLeft(3, '0');
            s2_t += s2.Semester.ToString().PadLeft(1, '0');
            return s1_t.CompareTo(s2_t);
        }

        /// <summary>
        /// 背景完成
        /// </summary>
        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Document inResult = (Document)e.Result;
            btnSave.Enabled = true;

            try
            {
                if (PrintSaveFile)
                {
                    if (cbIsDoc.Checked)
                    {
                        FolderBrowserDialog fbd = new FolderBrowserDialog();
                        fbd.Description = "請選擇訓導記錄表檔案儲存位置\n規格為(學號_身分證號_班級_座號_姓名)";
                        DialogResult dr = fbd.ShowDialog();
                        if (dr == System.Windows.Forms.DialogResult.OK)
                        {
                            foreach (StudentDataObj student in StudentSaveDic.Keys)
                            {
                                Document doc = StudentSaveDic[student];
                                StringBuilder sb = new StringBuilder();
                                sb.Append(fbd.SelectedPath + "\\");
                                sb.Append(student.StudentRecord.StudentNumber + "_");
                                sb.Append(student.StudentRecord.IDNumber + "_");
                                sb.Append((student.StudentRecord.Class != null ? student.StudentRecord.Class.Name : "") + "_");
                                sb.Append((student.StudentRecord.SeatNo.HasValue ? "" + student.StudentRecord.SeatNo.Value : "") + "_");
                                sb.Append(student.StudentRecord.Name + ".doc");

                                doc.Save(sb.ToString());
                            }
                            MsgBox.Show("學生訓導記錄表,列印完成!!");
                            System.Diagnostics.Process.Start("explorer", fbd.SelectedPath);
                        }
                        else
                        {
                            MsgBox.Show("已取消存檔!!");
                            return;
                        }
                    }
                    if (cbIsPdf.Checked)
                    {
                        FolderBrowserDialog fbd = new FolderBrowserDialog();
                        fbd.Description = "請選擇訓導記錄表檔案儲存位置\n規格為(學號_身分證號_班級_座號_姓名)";
                        DialogResult dr = fbd.ShowDialog();
                        if (dr == System.Windows.Forms.DialogResult.OK)
                        {
                            foreach (StudentDataObj student in StudentSaveDic.Keys)
                            {
                                Document doc = StudentSaveDic[student];
                                StringBuilder sb = new StringBuilder();
                                sb.Append(fbd.SelectedPath + "\\");
                                sb.Append(student.StudentRecord.StudentNumber + "_");
                                sb.Append(student.StudentRecord.IDNumber + "_");
                                sb.Append((student.StudentRecord.Class != null ? student.StudentRecord.Class.Name : "") + "_");
                                sb.Append((student.StudentRecord.SeatNo.HasValue ? "" + student.StudentRecord.SeatNo.Value : "") + "_");
                                sb.Append(student.StudentRecord.Name + ".pdf");


                                MemoryStream stream = new MemoryStream();

                                doc.Save(stream, SaveFormat.Doc);

                                Aspose.IO.Tools.SavePDFtoLocal(stream, sb.ToString());

                            }
                            MsgBox.Show("學生訓導記錄表,列印完成!!");
                            System.Diagnostics.Process.Start("explorer", fbd.SelectedPath);
                        }
                        else
                        {
                            MsgBox.Show("已取消存檔!!");
                            return;
                        }
                    }
                }
                else
                {
                    if (cbIsDoc.Checked)
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                        SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        SaveFileDialog1.FileName = string.Format("學生訓導紀錄表(高雄) {0}", DateTime.Now.ToString("yyyy-MM-dd HH-mm"));
                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            inResult.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                            MotherForm.SetStatusBarMessage("學生訓導記錄表,列印完成!!");
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("已取消存檔");
                            return;
                        }
                    }
                    if (cbIsPdf.Checked)
                    {
                        SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                        SaveFileDialog1.Filter = "Pdf Files|*.pdf";
                        SaveFileDialog1.FileName = string.Format("學生訓導紀錄表(高雄) {0}", DateTime.Now.ToString("yyyy-MM-dd HH-mm"));
                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            MemoryStream stream = new MemoryStream();

                            inResult.Save(stream, SaveFormat.Doc);

                            Aspose.IO.Tools.SavePDFtoLocal(stream, SaveFileDialog1.FileName);

                            Process.Start(SaveFileDialog1.FileName);
                            MotherForm.SetStatusBarMessage("學生訓導記錄表,列印完成!!");
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("已取消存檔");
                            return;
                        }
                    }
                }
            }
            catch
            {
                FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
            }

        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        private void Write(Cell cell, string text)
        {
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = SurrogatePairString(text);
            _run.Font.Size = 10;
            _run.Font.Name = "標楷體";
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
        }

        /// <summary>
        /// 以Cell為基準,向右移一格
        /// </summary>
        private Cell GetMoveRightCell(Cell cell, int count)
        {
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// 日常生活表現設定值
        /// </summary>
        private void GetBehaviorConfig()
        {
            DLBList1.Clear();
            DLBList2.Clear();

            TieDic1.Clear();

            K12.Data.Configuration.ConfigData cd = K12.Data.School.Configuration["DLBehaviorConfig"];


            if (!string.IsNullOrEmpty(cd["DailyBehavior"]))
            {
                XmlElement dailyBehavior = XmlHelper.LoadXml(cd["DailyBehavior"]);
                foreach (XmlElement item in dailyBehavior.SelectNodes("Item"))
                {

                    DLBList1.Add(item.GetAttribute("Name"));
                }

                TieDic1.Add("日常行為表現", dailyBehavior.GetAttribute("Name"));
            }

            if (!string.IsNullOrEmpty(cd["DailyLifeRecommend"]))
            {
                XmlElement dailyLifeRecommend = XmlHelper.LoadXml(cd["DailyLifeRecommend"]);
                DLBList2.Add("DailyLifeRecommend");

                TieDic1.Add("具體建議", dailyLifeRecommend.GetAttribute("Name"));
            }


            if (!string.IsNullOrEmpty(cd["GroupActivity"]))
            {
                XmlElement groupActivity = XmlHelper.LoadXml(cd["GroupActivity"]);
                DLBList2.Add("GroupActivity");

                TieDic1.Add("團體活動表現", groupActivity.GetAttribute("Name"));
            }


            if (!string.IsNullOrEmpty(cd["PublicService"]))
            {
                XmlElement publicService = XmlHelper.LoadXml(cd["PublicService"]);
                DLBList2.Add("PublicService");

                TieDic1.Add("公共服務表現", publicService.GetAttribute("Name"));
            }


            if (!string.IsNullOrEmpty(cd["SchoolSpecial"]))
            {
                XmlElement schoolSpecial = XmlHelper.LoadXml(cd["SchoolSpecial"]);
                DLBList2.Add("SchoolSpecial");

                TieDic1.Add("校內外特殊", schoolSpecial.GetAttribute("Name"));
            }
        }



        /// <summary>
        /// 建立已知資料
        /// </summary>
        private void SetNameIndex()
        {
            DicSummaryIndex.Clear();
            DicSummaryIndex.Add("事假一般", 2);
            DicSummaryIndex.Add("事假集會", 3);
            DicSummaryIndex.Add("病假一般", 4);
            DicSummaryIndex.Add("病假集會", 5);
            DicSummaryIndex.Add("曠課一般", 6);
            DicSummaryIndex.Add("曠課集會", 7);
            DicSummaryIndex.Add("公假一般", 8);
            DicSummaryIndex.Add("公假集會", 9);
            DicSummaryIndex.Add("喪假一般", 10);
            DicSummaryIndex.Add("喪假集會", 11);
            DicSummaryIndex.Add("大功", 12);
            DicSummaryIndex.Add("小功", 13);
            DicSummaryIndex.Add("嘉獎", 14);
            DicSummaryIndex.Add("大過", 15);
            DicSummaryIndex.Add("小過", 16);
            DicSummaryIndex.Add("警告", 17);

            UpdateCoddic.Clear();
            UpdateCoddic.Add("1", "新生");
            UpdateCoddic.Add("2", "畢業");
            UpdateCoddic.Add("3", "轉入");
            UpdateCoddic.Add("4", "轉出");
            UpdateCoddic.Add("5", "休學");
            UpdateCoddic.Add("6", "復學");
            UpdateCoddic.Add("7", "中輟");
            UpdateCoddic.Add("8", "續讀");
            UpdateCoddic.Add("9", "更正學籍");
        }

        /// <summary>
        /// 取得定義的統計資料Index
        /// </summary>
        private int GetSummaryIndex(string AttName)
        {
            if (DicSummaryIndex.ContainsKey(AttName))
            {
                return DicSummaryIndex[AttName];
            }
            else
            {
                return 0;
            }

        }

        /// <summary>
        /// 傳入異動代碼,取得異動原因
        /// </summary>
        private string GetUpdateRecordCode(string UpdateCode)
        {
            if (UpdateCoddic.ContainsKey(UpdateCode))
            {
                return UpdateCoddic[UpdateCode];
            }
            else
            {
                return "";
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void cbIsPdf_CheckedChanged(object sender, EventArgs e)
        {
            if (cbIsPdf.Checked)
            {
                cbIsDoc.Checked = false;
            }
            else
            {
                cbIsDoc.Checked = true;
            }
        }

        private void cbIsDoc_CheckedChanged(object sender, EventArgs e)
        {
            if (cbIsDoc.Checked)
            {
                cbIsPdf.Checked = false;

            }
            else
            {
                cbIsPdf.Checked = true;
            }
        }
    }

    public class SemesterSLR
    {
        public SemesterSLR(SLRecord slr)
        {
            SchoolYear = slr.SchoolYear;
            Semester = slr.Semester;
        }
        public int SchoolYear { get; set; }
        public int Semester { get; set; }
    }



}
