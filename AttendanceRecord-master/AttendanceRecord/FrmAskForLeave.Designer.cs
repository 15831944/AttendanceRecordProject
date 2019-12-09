namespace AttendanceRecord
{
    public partial class FrmAskForLeave
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblName = new System.Windows.Forms.Label();
            this.tbName = new System.Windows.Forms.TextBox();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cmStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.delByNOToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.timerClsResult = new System.Windows.Forms.Timer(this.components);
            this.dtPicker = new System.Windows.Forms.DateTimePicker();
            this.lblStartTime = new System.Windows.Forms.Label();
            this.lblEndTime = new System.Windows.Forms.Label();
            this.tbNO = new System.Windows.Forms.TextBox();
            this.lblNO = new System.Windows.Forms.Label();
            this.timeStartPicker = new System.Windows.Forms.DateTimePicker();
            this.timeEndPicker = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.lblStartDate = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.cmStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblName.Location = new System.Drawing.Point(123, 122);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(114, 33);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "姓名：";
            // 
            // tbName
            // 
            this.tbName.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbName.Location = new System.Drawing.Point(229, 114);
            this.tbName.Name = "tbName";
            this.tbName.Size = new System.Drawing.Size(279, 41);
            this.tbName.TabIndex = 3;
            this.tbName.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbName_KeyPress);
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.ContextMenuStrip = this.cmStrip;
            this.dgv.Location = new System.Drawing.Point(35, 417);
            this.dgv.MultiSelect = false;
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv.Size = new System.Drawing.Size(1292, 309);
            this.dgv.TabIndex = 8;
            // 
            // cmStrip
            // 
            this.cmStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.delByNOToolStripMenuItem});
            this.cmStrip.Name = "cmStrip";
            this.cmStrip.Size = new System.Drawing.Size(176, 34);
            // 
            // delByNOToolStripMenuItem
            // 
            this.delByNOToolStripMenuItem.Font = new System.Drawing.Font("微软雅黑", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.delByNOToolStripMenuItem.Name = "delByNOToolStripMenuItem";
            this.delByNOToolStripMenuItem.Size = new System.Drawing.Size(175, 30);
            this.delByNOToolStripMenuItem.Text = "删除该假条";
            this.delByNOToolStripMenuItem.Click += new System.EventHandler(this.delByNOToolStripMenuItem_Click);
            // 
            // btnSubmit
            // 
            this.btnSubmit.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSubmit.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnSubmit.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSubmit.Location = new System.Drawing.Point(1210, 347);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(108, 33);
            this.btnSubmit.TabIndex = 12;
            this.btnSubmit.Text = "提交";
            this.btnSubmit.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnSubmit.UseVisualStyleBackColor = false;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblResult.Location = new System.Drawing.Point(36, 746);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1291, 52);
            this.lblResult.TabIndex = 13;
            this.lblResult.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // timerClsResult
            // 
            this.timerClsResult.Interval = 5000;
            this.timerClsResult.Tick += new System.EventHandler(this.timerClsResult_Tick);
            // 
            // dtPicker
            // 
            this.dtPicker.CustomFormat = "yyyy-MM-dd";
            this.dtPicker.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtPicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPicker.Location = new System.Drawing.Point(229, 211);
            this.dtPicker.Name = "dtPicker";
            this.dtPicker.Size = new System.Drawing.Size(131, 29);
            this.dtPicker.TabIndex = 15;
            this.dtPicker.ValueChanged += new System.EventHandler(this.dtStartPicker_ValueChanged);
            // 
            // lblStartTime
            // 
            this.lblStartTime.AutoSize = true;
            this.lblStartTime.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblStartTime.Location = new System.Drawing.Point(57, 278);
            this.lblStartTime.Name = "lblStartTime";
            this.lblStartTime.Size = new System.Drawing.Size(180, 33);
            this.lblStartTime.TabIndex = 16;
            this.lblStartTime.Text = "起始时间：";
            // 
            // lblEndTime
            // 
            this.lblEndTime.AutoSize = true;
            this.lblEndTime.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblEndTime.Location = new System.Drawing.Point(646, 293);
            this.lblEndTime.Name = "lblEndTime";
            this.lblEndTime.Size = new System.Drawing.Size(180, 33);
            this.lblEndTime.TabIndex = 18;
            this.lblEndTime.Text = "终止时间：";
            // 
            // tbNO
            // 
            this.tbNO.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tbNO.Location = new System.Drawing.Point(824, 114);
            this.tbNO.Name = "tbNO";
            this.tbNO.ReadOnly = true;
            this.tbNO.Size = new System.Drawing.Size(259, 41);
            this.tbNO.TabIndex = 19;
            // 
            // lblNO
            // 
            this.lblNO.AutoSize = true;
            this.lblNO.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblNO.Location = new System.Drawing.Point(704, 122);
            this.lblNO.Name = "lblNO";
            this.lblNO.Size = new System.Drawing.Size(114, 33);
            this.lblNO.TabIndex = 20;
            this.lblNO.Text = "单号：";
            // 
            // timeStartPicker
            // 
            this.timeStartPicker.CustomFormat = "hh:mm:ss";
            this.timeStartPicker.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.timeStartPicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.timeStartPicker.Location = new System.Drawing.Point(229, 283);
            this.timeStartPicker.Name = "timeStartPicker";
            this.timeStartPicker.ShowUpDown = true;
            this.timeStartPicker.Size = new System.Drawing.Size(131, 29);
            this.timeStartPicker.TabIndex = 21;
            this.timeStartPicker.ValueChanged += new System.EventHandler(this.timeStartPicker_ValueChanged);
            // 
            // timeEndPicker
            // 
            this.timeEndPicker.CustomFormat = "hh:mm:ss";
            this.timeEndPicker.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.timeEndPicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.timeEndPicker.Location = new System.Drawing.Point(836, 298);
            this.timeEndPicker.Name = "timeEndPicker";
            this.timeEndPicker.ShowUpDown = true;
            this.timeEndPicker.Size = new System.Drawing.Size(124, 29);
            this.timeEndPicker.TabIndex = 22;
            this.timeEndPicker.ValueChanged += new System.EventHandler(this.timeEndPicker_ValueChanged);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(1143, 114);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(180, 94);
            this.label1.TabIndex = 23;
            this.label1.Text = "08:00 - 17:00 08:00 - 12:00 13:00 - 17:00";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblStartDate
            // 
            this.lblStartDate.AutoSize = true;
            this.lblStartDate.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblStartDate.Location = new System.Drawing.Point(57, 207);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(164, 33);
            this.lblStartDate.TabIndex = 24;
            this.lblStartDate.Text = "请假日期:";
            // 
            // FrmAskForLeave
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1360, 807);
            this.Controls.Add(this.lblStartDate);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.timeEndPicker);
            this.Controls.Add(this.timeStartPicker);
            this.Controls.Add(this.lblNO);
            this.Controls.Add(this.tbNO);
            this.Controls.Add(this.lblEndTime);
            this.Controls.Add(this.lblStartTime);
            this.Controls.Add(this.dtPicker);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.tbName);
            this.Controls.Add(this.lblName);
            this.MaximizeBox = false;
            this.Name = "FrmAskForLeave";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "请假";
            this.Load += new System.EventHandler(this.FrmAskForLeave_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.cmStrip.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.TextBox tbName;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerClsResult;
        private System.Windows.Forms.DateTimePicker dtPicker;
        private System.Windows.Forms.Label lblStartTime;
        private System.Windows.Forms.Label lblEndTime;
        private System.Windows.Forms.TextBox tbNO;
        private System.Windows.Forms.Label lblNO;
        private System.Windows.Forms.DateTimePicker timeStartPicker;
        private System.Windows.Forms.DateTimePicker timeEndPicker;
        private System.Windows.Forms.ContextMenuStrip cmStrip;
        private System.Windows.Forms.ToolStripMenuItem delByNOToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblStartDate;
    }
}

