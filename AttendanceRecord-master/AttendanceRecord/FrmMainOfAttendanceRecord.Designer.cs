namespace AttendanceRecord
{
    partial class FrmMainOfAttendanceRecord
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
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmMainOfAttendanceRecord));
            this.AttendanceRecordToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importAttendanceRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.setRestDayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.learningToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.AskForLeaveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.QuerySpecificDayARToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ARCalculateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lblUserName = new System.Windows.Forms.Label();
            this.nfiSystem = new System.Windows.Forms.NotifyIcon(this.components);
            this.cmStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.showMainFormToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.显示上传进度ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.文件FToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.新建NToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.打开OToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.保存SToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.另存为AToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.打印PToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.打印预览VToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.退出XToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.编辑EToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.撤消UToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.重复RToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.剪切TToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.复制CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.粘贴PToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.全选AToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.工具TToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.自定义CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.选项OToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.帮助HToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.内容CToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.索引IToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.搜索SToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
            this.关于AToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.GenerateWorkSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.mStrip = new System.Windows.Forms.MenuStrip();
            this.setTheDaysOfOverTimeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.SetTheDaysOfAskForLeaveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ExportARSummaryToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.FillUpARToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importARStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.cmStrip.SuspendLayout();
            this.mStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // AttendanceRecordToolStripMenuItem
            // 
            this.AttendanceRecordToolStripMenuItem.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.AttendanceRecordToolStripMenuItem.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.AttendanceRecordToolStripMenuItem.Name = "AttendanceRecordToolStripMenuItem";
            this.AttendanceRecordToolStripMenuItem.Size = new System.Drawing.Size(106, 25);
            this.AttendanceRecordToolStripMenuItem.Text = "考勤记录";
            // 
            // importAttendanceRToolStripMenuItem
            // 
            this.importAttendanceRToolStripMenuItem.Name = "importAttendanceRToolStripMenuItem";
            this.importAttendanceRToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.importAttendanceRToolStripMenuItem.Text = "导入";
            // 
            // setRestDayToolStripMenuItem
            // 
            this.setRestDayToolStripMenuItem.Name = "setRestDayToolStripMenuItem";
            this.setRestDayToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.setRestDayToolStripMenuItem.Text = "设定休息日";
            this.setRestDayToolStripMenuItem.Click += new System.EventHandler(this.importRestDayToolStripMenuItem_Click);
            // 
            // learningToolStripMenuItem
            // 
            this.learningToolStripMenuItem.BackColor = System.Drawing.SystemColors.Control;
            this.learningToolStripMenuItem.ForeColor = System.Drawing.Color.Gray;
            this.learningToolStripMenuItem.Name = "learningToolStripMenuItem";
            this.learningToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.learningToolStripMenuItem.Text = "早班 8:30";
            // 
            // AskForLeaveToolStripMenuItem
            // 
            this.AskForLeaveToolStripMenuItem.Name = "AskForLeaveToolStripMenuItem";
            this.AskForLeaveToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.AskForLeaveToolStripMenuItem.Text = "请假";
            this.AskForLeaveToolStripMenuItem.Click += new System.EventHandler(this.AskForLeaveToolStripMenuItem_Click);
            // 
            // QuerySpecificDayARToolStripMenuItem
            // 
            this.QuerySpecificDayARToolStripMenuItem.Name = "QuerySpecificDayARToolStripMenuItem";
            this.QuerySpecificDayARToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.QuerySpecificDayARToolStripMenuItem.Text = "查询当日出勤";
            // 
            // ARCalculateToolStripMenuItem
            // 
            this.ARCalculateToolStripMenuItem.Name = "ARCalculateToolStripMenuItem";
            this.ARCalculateToolStripMenuItem.Size = new System.Drawing.Size(206, 26);
            this.ARCalculateToolStripMenuItem.Text = "考勤记录计算";
            // 
            // lblUserName
            // 
            this.lblUserName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.lblUserName.BackColor = System.Drawing.SystemColors.Control;
            this.lblUserName.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblUserName.Location = new System.Drawing.Point(1108, 5);
            this.lblUserName.Name = "lblUserName";
            this.lblUserName.Size = new System.Drawing.Size(217, 19);
            this.lblUserName.TabIndex = 3;
            this.lblUserName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // nfiSystem
            // 
            this.nfiSystem.ContextMenuStrip = this.cmStrip;
            this.nfiSystem.Icon = ((System.Drawing.Icon)(resources.GetObject("nfiSystem.Icon")));
            this.nfiSystem.Visible = true;
            this.nfiSystem.Click += new System.EventHandler(this.nfiSystem_Click);
            this.nfiSystem.DoubleClick += new System.EventHandler(this.nfiSystem_DoubleClick);
            // 
            // cmStrip
            // 
            this.cmStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.showMainFormToolStripMenuItem,
            this.显示上传进度ToolStripMenuItem});
            this.cmStrip.Name = "cmStrip";
            this.cmStrip.Size = new System.Drawing.Size(149, 48);
            // 
            // showMainFormToolStripMenuItem
            // 
            this.showMainFormToolStripMenuItem.Name = "showMainFormToolStripMenuItem";
            this.showMainFormToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.showMainFormToolStripMenuItem.Text = "显示";
            this.showMainFormToolStripMenuItem.Click += new System.EventHandler(this.showMainFormToolStripMenuItem_Click);
            // 
            // 显示上传进度ToolStripMenuItem
            // 
            this.显示上传进度ToolStripMenuItem.Name = "显示上传进度ToolStripMenuItem";
            this.显示上传进度ToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.显示上传进度ToolStripMenuItem.Text = "显示上传进度";
            this.显示上传进度ToolStripMenuItem.Visible = false;
            // 
            // 文件FToolStripMenuItem
            // 
            this.文件FToolStripMenuItem.Name = "文件FToolStripMenuItem";
            this.文件FToolStripMenuItem.Size = new System.Drawing.Size(58, 21);
            this.文件FToolStripMenuItem.Text = "文件(&F)";
            // 
            // 新建NToolStripMenuItem
            // 
            this.新建NToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("新建NToolStripMenuItem.Image")));
            this.新建NToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.新建NToolStripMenuItem.Name = "新建NToolStripMenuItem";
            this.新建NToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.N)));
            this.新建NToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.新建NToolStripMenuItem.Text = "新建(&N)";
            // 
            // 打开OToolStripMenuItem
            // 
            this.打开OToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("打开OToolStripMenuItem.Image")));
            this.打开OToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.打开OToolStripMenuItem.Name = "打开OToolStripMenuItem";
            this.打开OToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.打开OToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.打开OToolStripMenuItem.Text = "打开(&O)";
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(162, 6);
            // 
            // 保存SToolStripMenuItem
            // 
            this.保存SToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("保存SToolStripMenuItem.Image")));
            this.保存SToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.保存SToolStripMenuItem.Name = "保存SToolStripMenuItem";
            this.保存SToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.保存SToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.保存SToolStripMenuItem.Text = "保存(&S)";
            // 
            // 另存为AToolStripMenuItem
            // 
            this.另存为AToolStripMenuItem.Name = "另存为AToolStripMenuItem";
            this.另存为AToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.另存为AToolStripMenuItem.Text = "另存为(&A)";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 6);
            // 
            // 打印PToolStripMenuItem
            // 
            this.打印PToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("打印PToolStripMenuItem.Image")));
            this.打印PToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.打印PToolStripMenuItem.Name = "打印PToolStripMenuItem";
            this.打印PToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.P)));
            this.打印PToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.打印PToolStripMenuItem.Text = "打印(&P)";
            // 
            // 打印预览VToolStripMenuItem
            // 
            this.打印预览VToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("打印预览VToolStripMenuItem.Image")));
            this.打印预览VToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.打印预览VToolStripMenuItem.Name = "打印预览VToolStripMenuItem";
            this.打印预览VToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.打印预览VToolStripMenuItem.Text = "打印预览(&V)";
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(162, 6);
            // 
            // 退出XToolStripMenuItem
            // 
            this.退出XToolStripMenuItem.Name = "退出XToolStripMenuItem";
            this.退出XToolStripMenuItem.Size = new System.Drawing.Size(165, 22);
            this.退出XToolStripMenuItem.Text = "退出(&X)";
            // 
            // 编辑EToolStripMenuItem
            // 
            this.编辑EToolStripMenuItem.Name = "编辑EToolStripMenuItem";
            this.编辑EToolStripMenuItem.Size = new System.Drawing.Size(59, 21);
            this.编辑EToolStripMenuItem.Text = "编辑(&E)";
            // 
            // 撤消UToolStripMenuItem
            // 
            this.撤消UToolStripMenuItem.Name = "撤消UToolStripMenuItem";
            this.撤消UToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Z)));
            this.撤消UToolStripMenuItem.Size = new System.Drawing.Size(161, 22);
            this.撤消UToolStripMenuItem.Text = "撤消(&U)";
            // 
            // 重复RToolStripMenuItem
            // 
            this.重复RToolStripMenuItem.Name = "重复RToolStripMenuItem";
            this.重复RToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 6);
            // 
            // 剪切TToolStripMenuItem
            // 
            this.剪切TToolStripMenuItem.Name = "剪切TToolStripMenuItem";
            this.剪切TToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 复制CToolStripMenuItem
            // 
            this.复制CToolStripMenuItem.Name = "复制CToolStripMenuItem";
            this.复制CToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 粘贴PToolStripMenuItem
            // 
            this.粘贴PToolStripMenuItem.Name = "粘贴PToolStripMenuItem";
            this.粘贴PToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 6);
            // 
            // 全选AToolStripMenuItem
            // 
            this.全选AToolStripMenuItem.Name = "全选AToolStripMenuItem";
            this.全选AToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 工具TToolStripMenuItem
            // 
            this.工具TToolStripMenuItem.Name = "工具TToolStripMenuItem";
            this.工具TToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 自定义CToolStripMenuItem
            // 
            this.自定义CToolStripMenuItem.Name = "自定义CToolStripMenuItem";
            this.自定义CToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 选项OToolStripMenuItem
            // 
            this.选项OToolStripMenuItem.Name = "选项OToolStripMenuItem";
            this.选项OToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 帮助HToolStripMenuItem
            // 
            this.帮助HToolStripMenuItem.Name = "帮助HToolStripMenuItem";
            this.帮助HToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 内容CToolStripMenuItem
            // 
            this.内容CToolStripMenuItem.Name = "内容CToolStripMenuItem";
            this.内容CToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 索引IToolStripMenuItem
            // 
            this.索引IToolStripMenuItem.Name = "索引IToolStripMenuItem";
            this.索引IToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // 搜索SToolStripMenuItem
            // 
            this.搜索SToolStripMenuItem.Name = "搜索SToolStripMenuItem";
            this.搜索SToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripSeparator5
            // 
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new System.Drawing.Size(6, 6);
            // 
            // 关于AToolStripMenuItem
            // 
            this.关于AToolStripMenuItem.Name = "关于AToolStripMenuItem";
            this.关于AToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripMenuItem
            // 
            this.toolStripMenuItem.Name = "toolStripMenuItem";
            this.toolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(32, 19);
            // 
            // GenerateWorkSummaryToolStripMenuItem
            // 
            this.GenerateWorkSummaryToolStripMenuItem.Name = "GenerateWorkSummaryToolStripMenuItem";
            this.GenerateWorkSummaryToolStripMenuItem.Size = new System.Drawing.Size(269, 26);
            this.GenerateWorkSummaryToolStripMenuItem.Text = "生成工作概况汇总表";
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(32, 19);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(32, 19);
            // 
            // mStrip
            // 
            this.mStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.importARStripMenuItem1,
            this.setTheDaysOfOverTimeToolStripMenuItem,
            this.SetTheDaysOfAskForLeaveToolStripMenuItem,
            this.ExportARSummaryToolStripMenuItem});
            this.mStrip.Location = new System.Drawing.Point(0, 0);
            this.mStrip.Name = "mStrip";
            this.mStrip.Size = new System.Drawing.Size(1337, 29);
            this.mStrip.TabIndex = 7;
            this.mStrip.Text = "menuStrip1";
            // 
            // setTheDaysOfOverTimeToolStripMenuItem
            // 
            this.setTheDaysOfOverTimeToolStripMenuItem.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.setTheDaysOfOverTimeToolStripMenuItem.Name = "setTheDaysOfOverTimeToolStripMenuItem";
            this.setTheDaysOfOverTimeToolStripMenuItem.Size = new System.Drawing.Size(127, 25);
            this.setTheDaysOfOverTimeToolStripMenuItem.Text = "设定加班日";
            this.setTheDaysOfOverTimeToolStripMenuItem.Click += new System.EventHandler(this.setTheDaysOfOverTimeToolStripMenuItem_Click);
            // 
            // SetTheDaysOfAskForLeaveToolStripMenuItem
            // 
            this.SetTheDaysOfAskForLeaveToolStripMenuItem.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SetTheDaysOfAskForLeaveToolStripMenuItem.Name = "SetTheDaysOfAskForLeaveToolStripMenuItem";
            this.SetTheDaysOfAskForLeaveToolStripMenuItem.Size = new System.Drawing.Size(148, 25);
            this.SetTheDaysOfAskForLeaveToolStripMenuItem.Text = "设置请假记录";
            this.SetTheDaysOfAskForLeaveToolStripMenuItem.Click += new System.EventHandler(this.SetTheDaysOfAskForLeaveToolStripMenuItem_Click);
            // 
            // ExportARSummaryToolStripMenuItem
            // 
            this.ExportARSummaryToolStripMenuItem.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ExportARSummaryToolStripMenuItem.Name = "ExportARSummaryToolStripMenuItem";
            this.ExportARSummaryToolStripMenuItem.Size = new System.Drawing.Size(224, 25);
            this.ExportARSummaryToolStripMenuItem.Text = "导出考勤汇总Excel档";
            this.ExportARSummaryToolStripMenuItem.Click += new System.EventHandler(this.ExportARSummaryToolStripMenuItem_Click);
            // 
            // FillUpARToolStripMenuItem
            // 
            this.FillUpARToolStripMenuItem.Name = "FillUpARToolStripMenuItem";
            this.FillUpARToolStripMenuItem.Size = new System.Drawing.Size(180, 26);
            this.FillUpARToolStripMenuItem.Text = "补卡";
            this.FillUpARToolStripMenuItem.Visible = false;
            this.FillUpARToolStripMenuItem.Click += new System.EventHandler(this.FillUpARToolStripMenuItem_Click);
            // 
            // importARStripMenuItem1
            // 
            this.importARStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.FillUpARToolStripMenuItem});
            this.importARStripMenuItem1.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.importARStripMenuItem1.Name = "importARStripMenuItem1";
            this.importARStripMenuItem1.Size = new System.Drawing.Size(182, 25);
            this.importARStripMenuItem1.Text = "导入考勤Excel档";
            this.importARStripMenuItem1.Click += new System.EventHandler(this.importARToolStripMenuItem_Click);
            // 
            // FrmMainOfAttendanceRecord
            // 
            this.ClientSize = new System.Drawing.Size(1337, 685);
            this.Controls.Add(this.mStrip);
            this.Controls.Add(this.lblUserName);
            this.IsMdiContainer = true;
            this.MaximizeBox = false;
            this.Name = "FrmMainOfAttendanceRecord";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "考勤记录汇总";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmMainOfAttendanceRecord_FormClosed);
            this.Load += new System.EventHandler(this.FrmMainOfAttendanceRecord_Load);
            this.cmStrip.ResumeLayout(false);
            this.mStrip.ResumeLayout(false);
            this.mStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.ToolStripMenuItem AttendanceRecordToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importAttendanceRToolStripMenuItem;
        private System.Windows.Forms.Label lblUserName;
        private System.Windows.Forms.ToolStripMenuItem AskForLeaveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem QuerySpecificDayARToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ARCalculateToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem setRestDayToolStripMenuItem;
        public System.Windows.Forms.NotifyIcon nfiSystem;
        private System.Windows.Forms.ContextMenuStrip cmStrip;
        private System.Windows.Forms.ToolStripMenuItem showMainFormToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem learningToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 文件FToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 新建NToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 打开OToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem 保存SToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 另存为AToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem 打印PToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 打印预览VToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem 退出XToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 编辑EToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 撤消UToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 重复RToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripMenuItem 剪切TToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 复制CToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 粘贴PToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripMenuItem 全选AToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 工具TToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 自定义CToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 选项OToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 帮助HToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 内容CToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 索引IToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 搜索SToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
        private System.Windows.Forms.ToolStripMenuItem 关于AToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem GenerateWorkSummaryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.MenuStrip mStrip;
        private System.Windows.Forms.ToolStripMenuItem 显示上传进度ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem setTheDaysOfOverTimeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem SetTheDaysOfAskForLeaveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ExportARSummaryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importARStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem FillUpARToolStripMenuItem;
        #endregion

        /*
        private System.Windows.Forms.MenuStrip mStrip;
        private System.Windows.Forms.ToolStripMenuItem QuerySalariesTSMI;
        private System.Windows.Forms.ToolStripMenuItem QuerySalariesByDepartmentTSMI;
        private System.Windows.Forms.ToolStripMenuItem QuerySalariesByNameToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem QueryQuantitiesByNameToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ProductsCostToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ProductsCostSummaryToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ProducingInfoToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        */
    }
}

