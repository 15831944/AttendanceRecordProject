namespace AttendanceRecord
{
    partial class FrmFillUpTheARRecord
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
            this.BtnARQuery = new System.Windows.Forms.Button();
            this.cbName = new System.Windows.Forms.ComboBox();
            this.lblName = new System.Windows.Forms.Label();
            this.dayPicker = new System.Windows.Forms.DateTimePicker();
            this.lblDay = new System.Windows.Forms.Label();
            this.lblTime = new System.Windows.Forms.Label();
            this.timerPicker = new System.Windows.Forms.DateTimePicker();
            this.btnFillUpRecord = new System.Windows.Forms.Button();
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
            // BtnARQuery
            // 
            this.BtnARQuery.Location = new System.Drawing.Point(1084, 194);
            this.BtnARQuery.Name = "BtnARQuery";
            this.BtnARQuery.Size = new System.Drawing.Size(212, 51);
            this.BtnARQuery.TabIndex = 11;
            this.BtnARQuery.Text = "查询";
            this.BtnARQuery.UseVisualStyleBackColor = true;
            this.BtnARQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // cbName
            // 
            this.cbName.Font = new System.Drawing.Font("宋体", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cbName.FormattingEnabled = true;
            this.cbName.Location = new System.Drawing.Point(175, 138);
            this.cbName.Name = "cbName";
            this.cbName.Size = new System.Drawing.Size(208, 32);
            this.cbName.TabIndex = 12;
            this.cbName.SelectedValueChanged += new System.EventHandler(this.cbName_SelectedValueChanged);
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
            // dayPicker
            // 
            this.dayPicker.CustomFormat = "yyyy-MM-dd";
            this.dayPicker.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dayPicker.Location = new System.Drawing.Point(728, 143);
            this.dayPicker.Name = "dayPicker";
            this.dayPicker.Size = new System.Drawing.Size(141, 29);
            this.dayPicker.TabIndex = 14;
            this.dayPicker.ValueChanged += new System.EventHandler(this.dayPicker_ValueChanged);
            // 
            // lblDay
            // 
            this.lblDay.Font = new System.Drawing.Font("宋体", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDay.Location = new System.Drawing.Point(629, 146);
            this.lblDay.Name = "lblDay";
            this.lblDay.Size = new System.Drawing.Size(93, 28);
            this.lblDay.TabIndex = 15;
            this.lblDay.Text = "日期：";
            // 
            // lblTime
            // 
            this.lblTime.Font = new System.Drawing.Font("宋体", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTime.Location = new System.Drawing.Point(893, 145);
            this.lblTime.Name = "lblTime";
            this.lblTime.Size = new System.Drawing.Size(93, 28);
            this.lblTime.TabIndex = 17;
            this.lblTime.Text = "时间：";
            // 
            // timerPicker
            // 
            this.timerPicker.CustomFormat = "hh:mm:ss";
            this.timerPicker.Format = System.Windows.Forms.DateTimePickerFormat.Time;
            this.timerPicker.Location = new System.Drawing.Point(992, 142);
            this.timerPicker.Name = "timerPicker";
            this.timerPicker.ShowUpDown = true;
            this.timerPicker.Size = new System.Drawing.Size(105, 29);
            this.timerPicker.TabIndex = 16;
            this.timerPicker.Value = new System.DateTime(2018, 10, 24, 8, 0, 0, 0);
            this.timerPicker.ValueChanged += new System.EventHandler(this.timerPicker_ValueChanged);
            // 
            // btnFillUpRecord
            // 
            this.btnFillUpRecord.Location = new System.Drawing.Point(1084, 271);
            this.btnFillUpRecord.Name = "btnFillUpRecord";
            this.btnFillUpRecord.Size = new System.Drawing.Size(212, 50);
            this.btnFillUpRecord.TabIndex = 18;
            this.btnFillUpRecord.Text = "补卡";
            this.btnFillUpRecord.UseVisualStyleBackColor = true;
            this.btnFillUpRecord.Click += new System.EventHandler(this.btnFillUpRecord_Click);
            // 
            // FrmFillUpTheARRecord
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(1334, 803);
            this.Controls.Add(this.btnFillUpRecord);
            this.Controls.Add(this.lblTime);
            this.Controls.Add(this.timerPicker);
            this.Controls.Add(this.lblDay);
            this.Controls.Add(this.dayPicker);
            this.Controls.Add(this.lblName);
            this.Controls.Add(this.cbName);
            this.Controls.Add(this.BtnARQuery);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.lblResult);
            this.Controls.Add(this.dgv);
            this.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(5);
            this.MaximizeBox = false;
            this.Name = "FrmFillUpTheARRecord";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "补卡";
            this.Load += new System.EventHandler(this.FrmFillUpTheARRecord_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.DataGridView dgv;
        private System.Windows.Forms.Label lblResult;
        private System.Windows.Forms.Timer timerRestoreTheLblResult;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.Button BtnARQuery;
        private System.Windows.Forms.ComboBox cbName;
        private System.Windows.Forms.Label lblName;
        private System.Windows.Forms.DateTimePicker dayPicker;
        private System.Windows.Forms.Label lblDay;
        private System.Windows.Forms.Label lblTime;
        private System.Windows.Forms.DateTimePicker timerPicker;
        private System.Windows.Forms.Button btnFillUpRecord;
    }
}