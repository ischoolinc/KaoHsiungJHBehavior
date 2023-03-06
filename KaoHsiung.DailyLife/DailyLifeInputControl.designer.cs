namespace KaoHsiung.DailyLife
{
    partial class DailyLifeInputControl
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgvTimes = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chGradeYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chStartTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chEndTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblSemester = new DevComponents.DotNetBar.LabelX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTimes)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvTimes
            // 
            this.dgvTimes.AllowUserToAddRows = false;
            this.dgvTimes.AllowUserToDeleteRows = false;
            this.dgvTimes.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvTimes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTimes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chGradeYear,
            this.chStartTime,
            this.chEndTime});
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle4.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvTimes.DefaultCellStyle = dataGridViewCellStyle4;
            this.dgvTimes.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvTimes.Location = new System.Drawing.Point(13, 38);
            this.dgvTimes.Name = "dgvTimes";
            this.dgvTimes.RowTemplate.Height = 24;
            this.dgvTimes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvTimes.Size = new System.Drawing.Size(462, 135);
            this.dgvTimes.TabIndex = 0;
            this.dgvTimes.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_CellClick);
            this.dgvTimes.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_CellEndEdit);
            this.dgvTimes.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgvTimes_CellValidating);
            this.dgvTimes.RowValidated += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_RowValidated);
            this.dgvTimes.RowValidating += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgvTimes_RowValidating);
            // 
            // chGradeYear
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.LightCyan;
            this.chGradeYear.DefaultCellStyle = dataGridViewCellStyle3;
            this.chGradeYear.HeaderText = "年級";
            this.chGradeYear.Name = "chGradeYear";
            this.chGradeYear.ReadOnly = true;
            // 
            // chStartTime
            // 
            this.chStartTime.HeaderText = "開始時間";
            this.chStartTime.Name = "chStartTime";
            this.chStartTime.Width = 150;
            // 
            // chEndTime
            // 
            this.chEndTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chEndTime.HeaderText = "截止時間";
            this.chEndTime.Name = "chEndTime";
            // 
            // lblSemester
            // 
            this.lblSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblSemester.BackgroundStyle.Class = "";
            this.lblSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblSemester.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.lblSemester.Location = new System.Drawing.Point(13, 9);
            this.lblSemester.Name = "lblSemester";
            this.lblSemester.Size = new System.Drawing.Size(142, 23);
            this.lblSemester.TabIndex = 1;
            this.lblSemester.Text = "112學年度　第１學期";
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnClose.Location = new System.Drawing.Point(400, 191);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "離開";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnSave.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnSave.Location = new System.Drawing.Point(319, 191);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.labelX1.Location = new System.Drawing.Point(15, 180);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(237, 39);
            this.labelX1.TabIndex = 3;
            this.labelX1.Text = "時間格式範例：\r\n輸入 3/6 會自動替換為2023/3/6 00:00";
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "年級";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "開始時間";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 150;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn3.HeaderText = "截止時間";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // DailyLifeInputControl
            // 
            this.AcceptButton = this.btnClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnSave;
            this.ClientSize = new System.Drawing.Size(489, 223);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.lblSemester);
            this.Controls.Add(this.dgvTimes);
            this.DoubleBuffered = true;
            this.Name = "DailyLifeInputControl";
            this.Text = "日常生活表現輸入時間設定";
            this.Load += new System.EventHandler(this.DailyLifeInputControl_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTimes)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgvTimes;
        private DevComponents.DotNetBar.LabelX lblSemester;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.DataGridViewTextBoxColumn chGradeYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn chStartTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn chEndTime;
    }
}