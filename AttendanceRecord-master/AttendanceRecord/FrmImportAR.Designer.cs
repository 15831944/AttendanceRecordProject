namespace AttendanceRecord
{
    partial class FrmImportAR
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
            this.lblResult = new System.Windows.Forms.Label();
            this.tb = new System.Windows.Forms.TextBox();
            this.btnImportEmpsInfo = new System.Windows.Forms.Button();
            this.timerRestoreTheLblResult = new System.Windows.Forms.Timer(this.components);
            this.btnViewTheUncertaiRecordInExcel = new System.Windows.Forms.Button();
            this.cbCheckSameNamesButDifferentMachineNo = new System.Windows.Forms.CheckBox();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.pb = new System.Windows.Forms.ProgressBar();
            this.dgv = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.SuspendLayout();
            // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblResult.Location = new System.Drawing.Point(49, 773);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1272, 38);
            this.lblResult.TabIndex = 15;
            // 
            // tb
            // 
            this.tb.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tb.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb.Location = new System.Drawing.Point(286, 98);
            this.tb.Name = "tb";
            this.tb.Size = new System.Drawing.Size(1032, 35);
            this.tb.TabIndex = 13;
            // 
            // btnImportEmpsInfo
            // 
            this.btnImportEmpsInfo.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnImportEmpsInfo.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportEmpsInfo.Location = new System.Drawing.Point(54, 81);
            this.btnImportEmpsInfo.Name = "btnImportEmpsInfo";
            this.btnImportEmpsInfo.Size = new System.Drawing.Size(196, 52);
            this.btnImportEmpsInfo.TabIndex = 12;
            this.btnImportEmpsInfo.Text = "导入考勤记录";
            this.btnImportEmpsInfo.UseVisualStyleBackColor = false;
            this.btnImportEmpsInfo.Click += new System.EventHandler(this.btnImportEmpsInfo_Click);
            // 
            // timerRestoreTheLblResult
            // 
            this.timerRestoreTheLblResult.Interval = 9000;
            this.timerRestoreTheLblResult.Tick += new System.EventHandler(this.timerRestoreTheLblResult_Tick);
            // 
            // btnViewTheUncertaiRecordInExcel
            // 
            this.btnViewTheUncertaiRecordInExcel.Enabled = false;
            this.btnViewTheUncertaiRecordInExcel.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnViewTheUncertaiRecordInExcel.Location = new System.Drawing.Point(1088, 174);
            this.btnViewTheUncertaiRecordInExcel.Name = "btnViewTheUncertaiRecordInExcel";
            this.btnViewTheUncertaiRecordInExcel.Size = new System.Drawing.Size(233, 38);
            this.btnViewTheUncertaiRecordInExcel.TabIndex = 19;
            this.btnViewTheUncertaiRecordInExcel.Text = "查看同名记录";
            this.btnViewTheUncertaiRecordInExcel.UseVisualStyleBackColor = true;
            this.btnViewTheUncertaiRecordInExcel.Click += new System.EventHandler(this.btnViewTheUncertaiRecordInExcel_Click);
            // 
            // cbCheckSameNamesButDifferentMachineNo
            // 
            this.cbCheckSameNamesButDifferentMachineNo.AutoSize = true;
            this.cbCheckSameNamesButDifferentMachineNo.Checked = true;
            this.cbCheckSameNamesButDifferentMachineNo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbCheckSameNamesButDifferentMachineNo.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbCheckSameNamesButDifferentMachineNo.Location = new System.Drawing.Point(744, 183);
            this.cbCheckSameNamesButDifferentMachineNo.Name = "cbCheckSameNamesButDifferentMachineNo";
            this.cbCheckSameNamesButDifferentMachineNo.Size = new System.Drawing.Size(323, 25);
            this.cbCheckSameNamesButDifferentMachineNo.TabIndex = 18;
            this.cbCheckSameNamesButDifferentMachineNo.Text = "检查姓名拼音相同或同名的用户";
            this.cbCheckSameNamesButDifferentMachineNo.UseVisualStyleBackColor = true;
            // 
            // lblPrompt
            // 
            this.lblPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPrompt.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblPrompt.Location = new System.Drawing.Point(54, 774);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(307, 37);
            this.lblPrompt.TabIndex = 17;
            this.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(360, 773);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(954, 37);
            this.pb.TabIndex = 16;
            this.pb.Visible = false;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(48, 230);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1273, 520);
            this.dgv.TabIndex = 14;
            // 
            // FrmImportAR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1374, 820);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.tb);
            this.Controls.Add(this.btnImportEmpsInfo);
            this.Controls.Add(this.btnViewTheUncertaiRecordInExcel);
            this.Controls.Add(this.cbCheckSameNamesButDifferentMachineNo);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.pb);
            this.Name = "FrmImportAR";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导入考勤记录";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.TextBox tb;
        private System.Windows.Forms.Button btnImportEmpsInfo;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.Button btnViewTheUncertaiRecordInExcel;
        private System.Windows.Forms.CheckBox cbCheckSameNamesButDifferentMachineNo;
        private System.Windows.Forms.Label lblPrompt;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.DataGridView dgv;
    }
}