namespace AttendanceRecord
{
    partial class FrmGenerateARSummary
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise; false.</param>
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
            this.btnGenerateWorkSchedule = new System.Windows.Forms.Button();
            this.mCalendar = new System.Windows.Forms.MonthCalendar();
            this.btnImportWorkSchedule = new System.Windows.Forms.Button();
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
            this.lblResult.Size = new System.Drawing.Size(1267, 49);
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
            this.pb.Location = new System.Drawing.Point(32, 742);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(1264, 49);
            this.pb.TabIndex = 6;
            this.pb.Visible = false;
            // 
            // btnGenerateWorkSchedule
            // 
            this.btnGenerateWorkSchedule.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnGenerateWorkSchedule.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnGenerateWorkSchedule.Location = new System.Drawing.Point(1104, 198);
            this.btnGenerateWorkSchedule.Name = "btnGenerateWorkSchedule";
            this.btnGenerateWorkSchedule.Size = new System.Drawing.Size(196, 52);
            this.btnGenerateWorkSchedule.TabIndex = 5;
            this.btnGenerateWorkSchedule.Text = "生成工作安排表";
            this.btnGenerateWorkSchedule.UseVisualStyleBackColor = false;
            this.btnGenerateWorkSchedule.Click += new System.EventHandler(this.btnGenerateWorkSchedule_Click);
            // 
            // mCalendar
            // 
            this.mCalendar.Location = new System.Drawing.Point(37, 36);
            this.mCalendar.Name = "mCalendar";
            this.mCalendar.TabIndex = 7;
            this.mCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.mCalendar_DateChanged);
            this.mCalendar.DateSelected += new System.Windows.Forms.DateRangeEventHandler(this.mCalendar_DateSelected);
            // 
            // btnImportWorkSchedule
            // 
            this.btnImportWorkSchedule.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnImportWorkSchedule.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportWorkSchedule.Location = new System.Drawing.Point(1104, 268);
            this.btnImportWorkSchedule.Name = "btnImportWorkSchedule";
            this.btnImportWorkSchedule.Size = new System.Drawing.Size(196, 52);
            this.btnImportWorkSchedule.TabIndex = 8;
            this.btnImportWorkSchedule.Text = "导入工作安排表";
            this.btnImportWorkSchedule.UseVisualStyleBackColor = false;
            this.btnImportWorkSchedule.Click += new System.EventHandler(this.btnImportWorkSchedule_Click);
            // 
            // FrmGenerateARSummary
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 803);
            this.Controls.Add(this.btnImportWorkSchedule);
            this.Controls.Add(this.mCalendar);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.btnGenerateWorkSchedule);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmGenerateARSummary";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "生成工作安排表";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Button btnGenerateWorkSchedule;
        private System.Windows.Forms.MonthCalendar mCalendar;
        private System.Windows.Forms.Button btnImportWorkSchedule;
    }
}