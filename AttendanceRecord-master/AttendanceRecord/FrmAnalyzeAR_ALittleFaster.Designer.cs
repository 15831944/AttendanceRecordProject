namespace AttendanceRecord
{
    partial class FrmAnalyzeAR_ALittleFaster
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
            this.mCalendar = new System.Windows.Forms.MonthCalendar();
            this.radioBtnToAWholePiece = new System.Windows.Forms.RadioButton();
            this.radioBtnSeparate = new System.Windows.Forms.RadioButton();
            this.btnExportARResult = new System.Windows.Forms.Button();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.pb = new System.Windows.Forms.ProgressBar();
            this.timerRestoreTheLblResult = new System.Windows.Forms.Timer(this.components);
            this.timerShowProgress = new System.Windows.Forms.Timer(this.components);
            this.timerCompleted = new System.Windows.Forms.Timer(this.components);
            this.lblResult = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // mCalendar
            // 
            this.mCalendar.Location = new System.Drawing.Point(56, 65);
            this.mCalendar.Name = "mCalendar";
            this.mCalendar.TabIndex = 8;
            this.mCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.mCalendar_DateChanged);
            // 
            // radioBtnToAWholePiece
            // 
            this.radioBtnToAWholePiece.AutoSize = true;
            this.radioBtnToAWholePiece.Checked = true;
            this.radioBtnToAWholePiece.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radioBtnToAWholePiece.Location = new System.Drawing.Point(954, 254);
            this.radioBtnToAWholePiece.Name = "radioBtnToAWholePiece";
            this.radioBtnToAWholePiece.Size = new System.Drawing.Size(217, 23);
            this.radioBtnToAWholePiece.TabIndex = 14;
            this.radioBtnToAWholePiece.TabStop = true;
            this.radioBtnToAWholePiece.Text = "导出至一张电子表格上";
            this.radioBtnToAWholePiece.UseVisualStyleBackColor = true;
            // 
            // radioBtnSeparate
            // 
            this.radioBtnSeparate.AutoSize = true;
            this.radioBtnSeparate.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.radioBtnSeparate.Location = new System.Drawing.Point(833, 254);
            this.radioBtnSeparate.Name = "radioBtnSeparate";
            this.radioBtnSeparate.Size = new System.Drawing.Size(103, 23);
            this.radioBtnSeparate.TabIndex = 13;
            this.radioBtnSeparate.Text = "分开导出";
            this.radioBtnSeparate.UseVisualStyleBackColor = true;
            // 
            // btnExportARResult
            // 
            this.btnExportARResult.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnExportARResult.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExportARResult.Location = new System.Drawing.Point(1192, 225);
            this.btnExportARResult.Name = "btnExportARResult";
            this.btnExportARResult.Size = new System.Drawing.Size(196, 52);
            this.btnExportARResult.TabIndex = 12;
            this.btnExportARResult.Text = "导出考勤记录";
            this.btnExportARResult.UseVisualStyleBackColor = false;
            this.btnExportARResult.Click += new System.EventHandler(this.btnExportARResult_Click);
            // 
            // dgv
            // 
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(21, 302);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1367, 381);
            this.dgv.TabIndex = 15;
            // 
            // lblPrompt
            // 
            this.lblPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPrompt.Location = new System.Drawing.Point(19, 701);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(218, 49);
            this.lblPrompt.TabIndex = 17;
            this.lblPrompt.Text = "sss";
            this.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblPrompt.Visible = false;
            // 
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(247, 701);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(1141, 49);
            this.pb.TabIndex = 16;
            this.pb.Visible = false;
            // 
            // timerRestoreTheLblResult
            // 
            this.timerRestoreTheLblResult.Interval = 7000;
            // 
            // timerShowProgress
            // 
            this.timerShowProgress.Interval = 10000;
            // 
            // timerCompleted
            // 
            this.timerCompleted.Interval = 513;
            // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Location = new System.Drawing.Point(21, 701);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1367, 49);
            this.lblResult.TabIndex = 18;
            this.lblResult.Visible = false;
            // 
            // FrmAnalyzeAR_ALittleFaster
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1409, 759);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.radioBtnToAWholePiece);
            this.Controls.Add(this.radioBtnSeparate);
            this.Controls.Add(this.btnExportARResult);
            this.Controls.Add(this.mCalendar);
            this.Name = "FrmAnalyzeAR_ALittleFaster";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "获取考勤记录_更快一点";
            this.Load += new System.EventHandler(this.FrmAnalyzeAR_ALittleFaster_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MonthCalendar mCalendar;
        private System.Windows.Forms.RadioButton radioBtnToAWholePiece;
        private System.Windows.Forms.RadioButton radioBtnSeparate;
        private System.Windows.Forms.Button btnExportARResult;
        private System.Windows.Forms.DataGridView dgv;
        public System.Windows.Forms.Label lblPrompt;
        public System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.Timer timerShowProgress;
        private System.Windows.Forms.Timer timerCompleted;
        private System.Windows.Forms.Label lblResult;
    }
}