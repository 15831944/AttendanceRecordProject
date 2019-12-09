namespace AttendanceRecord
{
    partial class FrmAnalyzeAR
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
            this.components = new System.ComponentModel.Container();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.lblResult = new System.Windows.Forms.Label();
            this.timerRestoreTheLblResult = new System.Windows.Forms.Timer(this.components);
            this.pb = new System.Windows.Forms.ProgressBar();
            this.btnExportARResult = new System.Windows.Forms.Button();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.mCalendar = new System.Windows.Forms.MonthCalendar();
            this.radioBtnSeparate = new System.Windows.Forms.RadioButton();
            this.radioBtnToAWholePiece = new System.Windows.Forms.RadioButton();
            this.timerShowProgress = new System.Windows.Forms.Timer(this.components);
            this.timerCompleted = new System.Windows.Forms.Timer(this.components);
            this.cb_Attendance_Machine = new System.Windows.Forms.ComboBox();
            this.lbl_Attendance_Machine = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(32, 338);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1267, 381);
            this.dgv.TabIndex = 2;
            // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Location = new System.Drawing.Point(32, 742);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1264, 49);
            this.lblResult.TabIndex = 3;
            this.lblResult.Click += new System.EventHandler(this.lblResult_Click);
            // 
            // timerRestoreTheLblResult
            // 
            this.timerRestoreTheLblResult.Interval = 7000;
            this.timerRestoreTheLblResult.Tick += new System.EventHandler(this.timerRestoreTheLblResult_Tick);
            // 
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(300, 742);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(998, 49);
            this.pb.TabIndex = 6;
            this.pb.Visible = false;
            // 
            // btnExportARResult
            // 
            this.btnExportARResult.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnExportARResult.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExportARResult.Location = new System.Drawing.Point(1103, 270);
            this.btnExportARResult.Name = "btnExportARResult";
            this.btnExportARResult.Size = new System.Drawing.Size(196, 52);
            this.btnExportARResult.TabIndex = 8;
            this.btnExportARResult.Text = "导出考勤记录";
            this.btnExportARResult.UseVisualStyleBackColor = false;
            this.btnExportARResult.Click += new System.EventHandler(this.btnExportARResult_Click);
            // 
            // lblPrompt
            // 
            this.lblPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPrompt.Location = new System.Drawing.Point(28, 742);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(264, 49);
            this.lblPrompt.TabIndex = 9;
            this.lblPrompt.Text = "sss";
            this.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblPrompt.Visible = false;
            // 
            // mCalendar
            // 
            this.mCalendar.Location = new System.Drawing.Point(36, 57);
            this.mCalendar.Name = "mCalendar";
            this.mCalendar.TabIndex = 7;
            this.mCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.mCalendar_DateChanged);
            // 
            // radioBtnSeparate
            // 
            this.radioBtnSeparate.AutoSize = true;
            this.radioBtnSeparate.Location = new System.Drawing.Point(829, 301);
            this.radioBtnSeparate.Name = "radioBtnSeparate";
            this.radioBtnSeparate.Size = new System.Drawing.Size(103, 23);
            this.radioBtnSeparate.TabIndex = 10;
            this.radioBtnSeparate.Text = "分开导出";
            this.radioBtnSeparate.UseVisualStyleBackColor = true;
            this.radioBtnSeparate.Visible = false;
            // 
            // radioBtnToAWholePiece
            // 
            this.radioBtnToAWholePiece.AutoSize = true;
            this.radioBtnToAWholePiece.Checked = true;
            this.radioBtnToAWholePiece.Location = new System.Drawing.Point(954, 301);
            this.radioBtnToAWholePiece.Name = "radioBtnToAWholePiece";
            this.radioBtnToAWholePiece.Size = new System.Drawing.Size(103, 23);
            this.radioBtnToAWholePiece.TabIndex = 11;
            this.radioBtnToAWholePiece.TabStop = true;
            this.radioBtnToAWholePiece.Text = "同名汇总";
            this.radioBtnToAWholePiece.UseVisualStyleBackColor = true;
            this.radioBtnToAWholePiece.Visible = false;
            // 
            // timerShowProgress
            // 
            this.timerShowProgress.Interval = 10000;
            this.timerShowProgress.Tick += new System.EventHandler(this.timerShowProgress_Tick);
            // 
            // timerCompleted
            // 
            this.timerCompleted.Interval = 513;
            this.timerCompleted.Tick += new System.EventHandler(this.timerCompleted_Tick);
            // 
            // cb_Attendance_Machine
            // 
            this.cb_Attendance_Machine.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Attendance_Machine.FormattingEnabled = true;
            this.cb_Attendance_Machine.Location = new System.Drawing.Point(502, 295);
            this.cb_Attendance_Machine.Name = "cb_Attendance_Machine";
            this.cb_Attendance_Machine.Size = new System.Drawing.Size(131, 27);
            this.cb_Attendance_Machine.TabIndex = 12;
            this.cb_Attendance_Machine.SelectedIndexChanged += new System.EventHandler(this.cb_Attendance_Machine_SelectedIndexChanged);
            // 
            // lbl_Attendance_Machine
            // 
            this.lbl_Attendance_Machine.AutoSize = true;
            this.lbl_Attendance_Machine.Location = new System.Drawing.Point(382, 301);
            this.lbl_Attendance_Machine.Name = "lbl_Attendance_Machine";
            this.lbl_Attendance_Machine.Size = new System.Drawing.Size(114, 19);
            this.lbl_Attendance_Machine.TabIndex = 13;
            this.lbl_Attendance_Machine.Text = "考勤机编号:";
            // 
            // FrmAnalyzeAR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 803);
            this.Controls.Add(this.lbl_Attendance_Machine);
            this.Controls.Add(this.cb_Attendance_Machine);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.radioBtnToAWholePiece);
            this.Controls.Add(this.radioBtnSeparate);
            this.Controls.Add(this.btnExportARResult);
            this.Controls.Add(this.mCalendar);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmAnalyzeAR";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "获取考勤记录";
            this.Load += new System.EventHandler(this.FrmAnalyzeAR_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.Button btnExportARResult;
        private System.Windows.Forms.MonthCalendar mCalendar;
        private System.Windows.Forms.RadioButton radioBtnSeparate;
        private System.Windows.Forms.RadioButton radioBtnToAWholePiece;
        private System.Windows.Forms.Timer timerShowProgress;
        public System.Windows.Forms.Label lblPrompt;
        public System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Timer timerCompleted;
        private System.Windows.Forms.ComboBox cb_Attendance_Machine;
        private System.Windows.Forms.Label lbl_Attendance_Machine;
    }
}