namespace AttendanceRecord
{
    partial class FrmQueryARByRange
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
            this.cM = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ExportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblResult = new System.Windows.Forms.Label();
            this.timerRestoreTheLblResult = new System.Windows.Forms.Timer(this.components);
            this.pb = new System.Windows.Forms.ProgressBar();
            this.dtStartDate = new System.Windows.Forms.DateTimePicker();
            this.dtEndDate = new System.Windows.Forms.DateTimePicker();
            this.lblStartDay = new System.Windows.Forms.Label();
            this.lblEndDay = new System.Windows.Forms.Label();
            this.btnQuery = new System.Windows.Forms.Button();
            this.cbName = new System.Windows.Forms.ComboBox();
            this.lblName = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.cM.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgv
            // 
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.ContextMenuStrip = this.cM;
            this.dgv.Location = new System.Drawing.Point(32, 338);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1267, 381);
            this.dgv.TabIndex = 2;
            // 
            // cM
            // 
            this.cM.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ExportToolStripMenuItem});
            this.cM.Name = "cM";
            this.cM.Size = new System.Drawing.Size(116, 34);
            // 
            // ExportToolStripMenuItem
            // 
            this.ExportToolStripMenuItem.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExportToolStripMenuItem.Name = "ExportToolStripMenuItem";
            this.ExportToolStripMenuItem.Size = new System.Drawing.Size(115, 30);
            this.ExportToolStripMenuItem.Text = "导出";
            this.ExportToolStripMenuItem.Click += new System.EventHandler(this.ExportToolStripMenuItem_Click);
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
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(32, 742);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(1264, 49);
            this.pb.TabIndex = 6;
            this.pb.Visible = false;
            // 
            // dtStartDate
            // 
            this.dtStartDate.CustomFormat = "yyyy年 M月d日";
            this.dtStartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtStartDate.Location = new System.Drawing.Point(658, 140);
            this.dtStartDate.Name = "dtStartDate";
            this.dtStartDate.Size = new System.Drawing.Size(200, 29);
            this.dtStartDate.TabIndex = 7;
            // 
            // dtEndDate
            // 
            this.dtEndDate.CustomFormat = "yyyy年 M月d日";
            this.dtEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtEndDate.Location = new System.Drawing.Point(1083, 141);
            this.dtEndDate.Name = "dtEndDate";
            this.dtEndDate.Size = new System.Drawing.Size(200, 29);
            this.dtEndDate.TabIndex = 8;
            // 
            // lblStartDay
            // 
            this.lblStartDay.Font = new System.Drawing.Font("宋体", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblStartDay.Location = new System.Drawing.Point(497, 141);
            this.lblStartDay.Name = "lblStartDay";
            this.lblStartDay.Size = new System.Drawing.Size(155, 28);
            this.lblStartDay.TabIndex = 9;
            this.lblStartDay.Text = "起始日期：";
            // 
            // lblEndDay
            // 
            this.lblEndDay.Font = new System.Drawing.Font("宋体", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblEndDay.Location = new System.Drawing.Point(928, 142);
            this.lblEndDay.Name = "lblEndDay";
            this.lblEndDay.Size = new System.Drawing.Size(155, 28);
            this.lblEndDay.TabIndex = 10;
            this.lblEndDay.Text = "终止日期：";
            // 
            // btnQuery
            // 
            this.btnQuery.Location = new System.Drawing.Point(1083, 263);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(212, 51);
            this.btnQuery.TabIndex = 11;
            this.btnQuery.Text = "查询";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // cbName
            // 
            this.cbName.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbName.FormattingEnabled = true;
            this.cbName.Location = new System.Drawing.Point(175, 138);
            this.cbName.Name = "cbName";
            this.cbName.Size = new System.Drawing.Size(208, 32);
            this.cbName.TabIndex = 12;
            this.cbName.TextChanged += new System.EventHandler(this.cbName_TextChanged);
            // 
            // lblName
            // 
            this.lblName.Font = new System.Drawing.Font("宋体", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblName.Location = new System.Drawing.Point(70, 142);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(99, 28);
            this.lblName.TabIndex = 13;
            this.lblName.Text = "姓名：";
            // 
            // FrmQueryARByRange
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 803);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.cbName);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.lblEndDay);
            this.Controls.Add(this.lblStartDay);
            this.Controls.Add(this.dtEndDate);
            this.Controls.Add(this.dtStartDate);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmQueryARByRange";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "查询某日期范围的考勤记录";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.cM.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.DateTimePicker dtStartDate;
        private System.Windows.Forms.DateTimePicker dtEndDate;
        private System.Windows.Forms.Label lblStartDay;
        private System.Windows.Forms.Label lblEndDay;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.ComboBox cbName;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.ContextMenuStrip cM;
        private System.Windows.Forms.ToolStripMenuItem ExportToolStripMenuItem;
    }
}