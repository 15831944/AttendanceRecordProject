namespace AttendanceRecord
{
    partial class FrmRestDay_Backup
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
            this.btnImportRestDay = new System.Windows.Forms.Button();
            this.tbRestDayPath = new System.Windows.Forms.TextBox();
            this.btnGenerateDefaultRestDays = new System.Windows.Forms.Button();
            this.mCalendar = new System.Windows.Forms.MonthCalendar();
            this.cms = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.delToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.cms.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(32, 397);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1267, 322);
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
            // 
            // timerRestoreTheLblResult
            // 
            this.timerRestoreTheLblResult.Interval = 7000;
            this.timerRestoreTheLblResult.Tick += new System.EventHandler(this.timerRestoreTheLblResult_Tick);
            // 
            // btnImportRestDay
            // 
            this.btnImportRestDay.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnImportRestDay.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportRestDay.Location = new System.Drawing.Point(32, 311);
            this.btnImportRestDay.Name = "btnImportRestDay";
            this.btnImportRestDay.Size = new System.Drawing.Size(196, 52);
            this.btnImportRestDay.TabIndex = 8;
            this.btnImportRestDay.Text = "导入休息日文档";
            this.btnImportRestDay.UseVisualStyleBackColor = false;
            this.btnImportRestDay.Click += new System.EventHandler(this.btnImportRestDay_Click);
            // 
            // tbRestDayPath
            // 
            this.tbRestDayPath.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbRestDayPath.Location = new System.Drawing.Point(252, 322);
            this.tbRestDayPath.Name = "tbRestDayPath";
            this.tbRestDayPath.Size = new System.Drawing.Size(803, 41);
            this.tbRestDayPath.TabIndex = 9;
            // 
            // btnGenerateDefaultRestDays
            // 
            this.btnGenerateDefaultRestDays.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnGenerateDefaultRestDays.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnGenerateDefaultRestDays.Location = new System.Drawing.Point(1103, 316);
            this.btnGenerateDefaultRestDays.Name = "btnGenerateDefaultRestDays";
            this.btnGenerateDefaultRestDays.Size = new System.Drawing.Size(196, 52);
            this.btnGenerateDefaultRestDays.TabIndex = 10;
            this.btnGenerateDefaultRestDays.Text = "生成默认的休息日";
            this.btnGenerateDefaultRestDays.UseVisualStyleBackColor = false;
            this.btnGenerateDefaultRestDays.Click += new System.EventHandler(this.btnGenerateDefaultRestDays_Click);
            // 
            // mCalendar
            // 
            this.mCalendar.Location = new System.Drawing.Point(32, 37);
            this.mCalendar.Name = "mCalendar";
            this.mCalendar.TabIndex = 11;
            this.mCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.mCalendar_DateChanged);
            // 
            // cms
            // 
            this.cms.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.delToolStripMenuItem});
            this.cms.Name = "cms";
            this.cms.Size = new System.Drawing.Size(101, 26);
            // 
            // delToolStripMenuItem
            // 
            this.delToolStripMenuItem.Name = "delToolStripMenuItem";
            this.delToolStripMenuItem.Size = new System.Drawing.Size(100, 22);
            this.delToolStripMenuItem.Text = "删除";
            this.delToolStripMenuItem.Click += new System.EventHandler(this.delToolStripMenuItem_Click);
            // 
            // FrmRestDay
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 803);
            this.Controls.Add(this.mCalendar);
            this.Controls.Add(this.btnGenerateDefaultRestDays);
            this.Controls.Add(this.tbRestDayPath);
            this.Controls.Add(this.btnImportRestDay);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmRestDay";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导入休息日文档";
            this.Load += new System.EventHandler(this.FrmRestDay_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.cms.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.Button btnImportRestDay;
        private System.Windows.Forms.TextBox tbRestDayPath;
        private System.Windows.Forms.Button btnGenerateDefaultRestDays;
        private System.Windows.Forms.MonthCalendar mCalendar;
        private System.Windows.Forms.ContextMenuStrip cms;
        private System.Windows.Forms.ToolStripMenuItem delToolStripMenuItem;
    }
}