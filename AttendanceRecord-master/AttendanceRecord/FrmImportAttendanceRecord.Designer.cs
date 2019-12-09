namespace AttendanceRecord
{
    partial class FrmImportAttendanceRecord
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
            this.btnImportEmpsInfo = new System.Windows.Forms.Button();
            this.tb = new System.Windows.Forms.TextBox();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.lblResult = new System.Windows.Forms.Label();
            this.timerRestoreTheLblResult = new System.Windows.Forms.Timer(this.components);
            this.pb = new System.Windows.Forms.ProgressBar();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.cbCheckSameNamesButDifferentMachineNo = new System.Windows.Forms.CheckBox();
            this.btnSwitch = new System.Windows.Forms.Button();
            this.dgv_same_name = new System.Windows.Forms.DataGridView();
            this.dgv_same_pinyin_of_name = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_same_name)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_same_pinyin_of_name)).BeginInit();
            this.SuspendLayout();
            // 
            // btnImportEmpsInfo
            // 
            this.btnImportEmpsInfo.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnImportEmpsInfo.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportEmpsInfo.Location = new System.Drawing.Point(33, 89);
            this.btnImportEmpsInfo.Name = "btnImportEmpsInfo";
            this.btnImportEmpsInfo.Size = new System.Drawing.Size(196, 52);
            this.btnImportEmpsInfo.TabIndex = 0;
            this.btnImportEmpsInfo.Text = "导入考勤记录";
            this.btnImportEmpsInfo.UseVisualStyleBackColor = false;
            this.btnImportEmpsInfo.Click += new System.EventHandler(this.btnImportEmpsARInfo_Click);
            // 
            // tb
            // 
            this.tb.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tb.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb.Location = new System.Drawing.Point(265, 106);
            this.tb.Name = "tb";
            this.tb.Size = new System.Drawing.Size(1032, 35);
            this.tb.TabIndex = 1;
            // 
            // dgv
            // 
            this.dgv.AllowUserToAddRows = false;
            this.dgv.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Location = new System.Drawing.Point(33, 230);
            this.dgv.Name = "dgv";
            this.dgv.RowTemplate.Height = 23;
            this.dgv.Size = new System.Drawing.Size(1267, 468);
            this.dgv.TabIndex = 2;
            this.dgv.Visible = false;
              // 
            // lblResult
            // 
            this.lblResult.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblResult.Location = new System.Drawing.Point(37, 725);
            this.lblResult.Name = "lblResult";
            this.lblResult.Size = new System.Drawing.Size(1263, 37);
            this.lblResult.TabIndex = 3;
            // 
            // timerRestoreTheLblResult
            // 
            this.timerRestoreTheLblResult.Interval = 9000;
            this.timerRestoreTheLblResult.Tick += new System.EventHandler(this.timerRestoreTheLblResult_Tick);
            // 
            // pb
            // 
            this.pb.Location = new System.Drawing.Point(346, 725);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(954, 37);
            this.pb.TabIndex = 6;
            this.pb.Visible = false;
            // 
            // lblPrompt
            // 
            this.lblPrompt.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblPrompt.Location = new System.Drawing.Point(33, 725);
            this.lblPrompt.Name = "lblPrompt";
            this.lblPrompt.Size = new System.Drawing.Size(307, 37);
            this.lblPrompt.TabIndex = 7;
            this.lblPrompt.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cbCheckSameNamesButDifferentMachineNo
            // 
            this.cbCheckSameNamesButDifferentMachineNo.AutoSize = true;
            this.cbCheckSameNamesButDifferentMachineNo.Checked = true;
            this.cbCheckSameNamesButDifferentMachineNo.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbCheckSameNamesButDifferentMachineNo.Location = new System.Drawing.Point(781, 191);
            this.cbCheckSameNamesButDifferentMachineNo.Name = "cbCheckSameNamesButDifferentMachineNo";
            this.cbCheckSameNamesButDifferentMachineNo.Size = new System.Drawing.Size(294, 23);
            this.cbCheckSameNamesButDifferentMachineNo.TabIndex = 8;
            this.cbCheckSameNamesButDifferentMachineNo.Text = "检查姓名拼音相同的或同名用户";
            this.cbCheckSameNamesButDifferentMachineNo.UseVisualStyleBackColor = true;
             // 
            // btnSwitch
            // 
            this.btnSwitch.Location = new System.Drawing.Point(1094, 186);
            this.btnSwitch.Name = "btnSwitch";
            this.btnSwitch.Size = new System.Drawing.Size(206, 28);
            this.btnSwitch.TabIndex = 9;
            this.btnSwitch.Text = "查看同名记录";
            this.btnSwitch.UseVisualStyleBackColor = true;
            this.btnSwitch.Click += new System.EventHandler(this.btnSwitch_Click);
            // 
            // dgv_same_name
            // 
            this.dgv_same_name.AllowUserToAddRows = false;
            this.dgv_same_name.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_same_name.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_same_name.Location = new System.Drawing.Point(33, 230);
            this.dgv_same_name.Name = "dgv_same_name";
            this.dgv_same_name.RowTemplate.Height = 23;
            this.dgv_same_name.Size = new System.Drawing.Size(1267, 468);
            this.dgv_same_name.TabIndex = 10;
            this.dgv_same_name.Visible = false;
            // 
            // dgv_same_pinyin_of_name
            // 
            this.dgv_same_pinyin_of_name.AllowUserToAddRows = false;
            this.dgv_same_pinyin_of_name.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv_same_pinyin_of_name.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv_same_pinyin_of_name.Location = new System.Drawing.Point(33, 230);
            this.dgv_same_pinyin_of_name.Name = "dgv_same_pinyin_of_name";
            this.dgv_same_pinyin_of_name.RowTemplate.Height = 23;
            this.dgv_same_pinyin_of_name.Size = new System.Drawing.Size(1267, 468);
            this.dgv_same_pinyin_of_name.TabIndex = 11;
            // 
            // FrmImportAttendanceRecord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 782);
            this.Controls.Add(this.dgv_same_pinyin_of_name);
            this.Controls.Add(this.dgv_same_name);
            this.Controls.Add(this.btnSwitch);
            this.Controls.Add(this.cbCheckSameNamesButDifferentMachineNo);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Controls.Add(this.tb);
            this.Controls.Add(this.btnImportEmpsInfo);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmImportAttendanceRecord";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导入考勤记录";
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_same_name)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv_same_pinyin_of_name)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnImportEmpsInfo;
        private System.Windows.Forms.TextBox tb;
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Label lblPrompt;
        private System.Windows.Forms.CheckBox cbCheckSameNamesButDifferentMachineNo;
        private System.Windows.Forms.Button btnSwitch;
        private System.Windows.Forms.DataGridView dgv_same_name;
        private System.Windows.Forms.DataGridView dgv_same_pinyin_of_name;
    }
}