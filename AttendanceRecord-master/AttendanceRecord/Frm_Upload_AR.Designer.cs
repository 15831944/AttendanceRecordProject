namespace AttendanceRecord
{
    partial class Frm_Upload_AR
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
            this.tb = new System.Windows.Forms.TextBox();
            this.btnImportEmpsInfo = new System.Windows.Forms.Button();
            this.timerReadProcessFromSvr = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // tb
            // 
            this.tb.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tb.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tb.Location = new System.Drawing.Point(305, 99);
            this.tb.Name = "tb";
            this.tb.Size = new System.Drawing.Size(1032, 35);
            this.tb.TabIndex = 3;
            // 
            // btnImportEmpsInfo
            // 
            this.btnImportEmpsInfo.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnImportEmpsInfo.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnImportEmpsInfo.Location = new System.Drawing.Point(73, 82);
            this.btnImportEmpsInfo.Name = "btnImportEmpsInfo";
            this.btnImportEmpsInfo.Size = new System.Drawing.Size(196, 52);
            this.btnImportEmpsInfo.TabIndex = 2;
            this.btnImportEmpsInfo.Text = "上传考勤记录";
            this.btnImportEmpsInfo.UseVisualStyleBackColor = false;
            this.btnImportEmpsInfo.Click += new System.EventHandler(this.btnImportEmpsInfo_Click);
            // 
            // timerReadProcessFromSvr
            // 
            this.timerReadProcessFromSvr.Tick += new System.EventHandler(this.timerReadProcessFromSvr_Tick);
            // 
            // Frm_Upload_AR
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1508, 719);
            this.Controls.Add(this.tb);
            this.Controls.Add(this.btnImportEmpsInfo);
            this.MaximizeBox = false;
            this.Name = "Frm_Upload_AR";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "上传考勤记录";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb;
        private System.Windows.Forms.Button btnImportEmpsInfo;
        private System.Windows.Forms.Timer timerReadProcessFromSvr;
    }
}