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
            this.dgv = new System.Windows.Forms.DataGridView();
            this.cmStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.delByNOToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSubmit = new System.Windows.Forms.Button();
            this.lblResult = new System.Windows.Forms.Label();
            this.timerClsResult = new System.Windows.Forms.Timer(this.components);
            this.dtPicker = new System.Windows.Forms.DateTimePicker();
            this.lblStartDate = new System.Windows.Forms.Label();
            this.cbName = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cbTimeSection = new System.Windows.Forms.ComboBox();
            this.dgvARDetail = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.cmStrip.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvARDetail)).BeginInit();
            this.SuspendLayout();
            // 
            // lblName
            // 
            this.lblName.AutoSize = true;
            this.lblName.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblName.Location = new System.Drawing.Point(111, 69);
            this.lblName.Name = "lblName";
            this.lblName.Size = new System.Drawing.Size(85, 24);
            this.lblName.TabIndex = 0;
            this.lblName.Text = "姓名：";
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.ContextMenuStrip = this.cmStrip;
            this.dgv.Location = new System.Drawing.Point(35, 599);
            this.dgv.MultiSelect = false;
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv.Size = new System.Drawing.Size(1292, 134);
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
            this.btnSubmit.Location = new System.Drawing.Point(1219, 187);
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
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblResult.Location = new System.Drawing.Point(36, 753);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1291, 45);
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
            this.dtPicker.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtPicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtPicker.Location = new System.Drawing.Point(214, 124);
            this.dtPicker.Name = "dtPicker";
            this.dtPicker.Size = new System.Drawing.Size(245, 35);
            this.dtPicker.TabIndex = 15;
            this.dtPicker.ValueChanged += new System.EventHandler(this.dtPicker_ValueChanged);
            // 
            // lblStartDate
            // 
            this.lblStartDate.AutoSize = true;
            this.lblStartDate.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblStartDate.Location = new System.Drawing.Point(56, 131);
            this.lblStartDate.Name = "lblStartDate";
            this.lblStartDate.Size = new System.Drawing.Size(123, 24);
            this.lblStartDate.TabIndex = 24;
            this.lblStartDate.Text = "请假日期:";
            // 
            // cbName
            // 
            this.cbName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbName.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbName.FormattingEnabled = true;
            this.cbName.Location = new System.Drawing.Point(214, 69);
            this.cbName.Name = "cbName";
            this.cbName.Size = new System.Drawing.Size(245, 32);
            this.cbName.TabIndex = 25;
            this.cbName.SelectedIndexChanged += new System.EventHandler(this.cbName_SelectedIndexChanged);
            this.cbName.SelectedValueChanged += new System.EventHandler(this.cbName_SelectedValueChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(31, 190);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(160, 24);
            this.label2.TabIndex = 26;
            this.label2.Text = "请假时间段：";
            // 
            // cbTimeSection
            // 
            this.cbTimeSection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbTimeSection.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbTimeSection.FormattingEnabled = true;
            this.cbTimeSection.Items.AddRange(new object[] {
            "08:00 - 12:00",
            "13:00 - 17:00",
            "08:00 - 17:00"});
            this.cbTimeSection.Location = new System.Drawing.Point(214, 187);
            this.cbTimeSection.Name = "cbTimeSection";
            this.cbTimeSection.Size = new System.Drawing.Size(245, 32);
            this.cbTimeSection.TabIndex = 27;
            // 
            // dgvARDetail
            // 
            this.dgvARDetail.AllowUserToAddRows = false;
            this.dgvARDetail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvARDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvARDetail.ContextMenuStrip = this.cmStrip;
            this.dgvARDetail.Location = new System.Drawing.Point(35, 270);
            this.dgvARDetail.MultiSelect = false;
            this.dgvARDetail.Name = "dgvARDetail";
            this.dgvARDetail.RowTemplate.Height = 23;
            this.dgvARDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvARDetail.Size = new System.Drawing.Size(1292, 309);
            this.dgvARDetail.TabIndex = 28;
            // 
            // FrmAskForLeave
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1360, 807);
            this.Controls.Add(this.dgvARDetail);
            this.Controls.Add(this.cbTimeSection);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cbName);
            this.Controls.Add(this.lblStartDate);
            this.Controls.Add(this.dtPicker);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.btnSubmit);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.lblName);
            this.MaximizeBox = false;
            this.Name = "FrmAskForLeave";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "请假";
            this.Load += new System.EventHandler(this.FrmAskForLeave_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.cmStrip.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvARDetail)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Button btnSubmit;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerClsResult;
        private System.Windows.Forms.DateTimePicker dtPicker;
        private System.Windows.Forms.ContextMenuStrip cmStrip;
        private System.Windows.Forms.ToolStripMenuItem delByNOToolStripMenuItem;
        private System.Windows.Forms.Label lblStartDate;
        private System.Windows.Forms.ComboBox cbName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbTimeSection;
        private System.Windows.Forms.DataGridView dgvARDetail;
    }
}

