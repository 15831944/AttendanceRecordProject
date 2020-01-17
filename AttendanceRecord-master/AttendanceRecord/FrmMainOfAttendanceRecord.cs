using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using AttendanceRecord.Action;
using AttendanceRecord.Entities;
using Tools;
namespace AttendanceRecord
{
    public partial class FrmMainOfAttendanceRecord : Form
    {
        #region Version Info
        //=====================================================================
        // Project Name        :    BaseDao  
        // Project Description : 
        // Class Name          :    Class1
        // File Name           :    Class1
        // Namespace           :    BaseDao 
        // Class Version       :    v1.0.0.0
        // Class Description   : 
        // CLR                 :    4.0.30319.42000  
        // Author              :    董   魁  (ccie20079@126.com)
        // Addr                :    中国  陕西 咸阳    
        // Create Time         :    2019-10-22 14:57:19
        // Modifier:     
        // Update Time         :    2019-10-22 14:57:19
        //======================================================================
        // Copyright © DGCZ  2019 . All rights reserved.
        // =====================================================================
        #endregion
        public FrmMainOfAttendanceRecord()
        {
            InitializeComponent();
        }
        private FrmImportAR frmImportAR = null;
        private FrmAnalyzeAR frmAnalyzeAR = null;
        private FrmQueryARByRange frmQueryARByRange = null;
        private FrmAskForLeave frmAskForLeave = null;
        private FrmTheDaysOfOvertime frmTheDaysOfOvertime = null;
        private FrmLearning frmLearning = null;
        private FrmFillUpTheARRecord frmFillUpTheARRecord = null;
        private Frm_Upload_AR frmUploadAR = null;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void FrmMainOfAttendanceRecord_Load(object sender, EventArgs e)
        {
            this.Text += " " + Application.ProductVersion;
            this.lblUserName.Text = Program._userInfo.Department + "   " + Program._userInfo.User_Name; ;
            this.lblUserName.BackColor = this.mStrip.BackColor;
            nfiSystem.Text = "杜克普考勤软件：" + Application.ProductVersion;
        }
     
        private void FrmMainOfAttendanceRecord_FormClosed(object sender, FormClosedEventArgs e)
        {
            FrmAnalyzeAR.killHwndOfXls();
            V_Work_Schedule.killHwndOfXls();

            AppManagement.closeAllExcel();
        }
          private void AskForLeaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
        }
        /// <summary>
        /// 导入休息日
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void importRestDayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(FrmTheDaysOfOvertime._action)) return;
            if (frmTheDaysOfOvertime == null || frmTheDaysOfOvertime.IsDisposed)
            {
                frmTheDaysOfOvertime = new FrmTheDaysOfOvertime();
                frmTheDaysOfOvertime.Show();
            }
            frmTheDaysOfOvertime.MdiParent = this;
            frmTheDaysOfOvertime.Activate();
        }

        private void nfiSystem_DoubleClick(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            this.nfiSystem.Visible = false;
            //显示正常图标。
        }

      
        private void showMainFormToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Maximized;
            this.ShowInTaskbar = true;
            this.nfiSystem.Visible = false;
            
        }
        private void nfiSystem_Click(object sender, EventArgs e)
        {
            if (((MouseEventArgs)e).Button == MouseButtons.Right) return;
            if (frmAnalyzeAR != null) {
                if (nfiSystem.Visible)
                {
                    nfiSystem.ShowBalloonTip(100000, "提示：", frmAnalyzeAR.lblPrompt.Text.Trim() + "   " + Math.Round((decimal)(frmAnalyzeAR.pb.Value / (decimal)frmAnalyzeAR.pb.Maximum) * 100, 2).ToString() + "%", ToolTipIcon.Info);
                }
            }
        }
        private void importARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(FrmImportAR._action)) return;
            if (frmImportAR == null || frmImportAR.IsDisposed)
            {
                frmImportAR = new FrmImportAR();
                frmImportAR.Show();
            }
            frmImportAR.MdiParent = this;
            frmImportAR.Activate();
        }

        private void halfPastEightToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(FrmLearning._action)) return;
            if (frmLearning == null || frmLearning.IsDisposed)
            {
                frmLearning = new FrmLearning();
                frmLearning.Show();
            }
            frmLearning.MdiParent = this;
            frmLearning.Activate();
        }

        private void QuerySpecificDayAR_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!ARActionHelper.getAction(FrmQueryARByRange._action)) return;
            if (frmQueryARByRange == null || frmQueryARByRange.IsDisposed)
            {
                frmQueryARByRange = new FrmQueryARByRange();
                frmQueryARByRange.Show();
            }
            frmQueryARByRange.MdiParent = this;
            frmQueryARByRange.Activate();
        }

        private void ARCalculate_ToolStripMenuItem_Click(object sender, EventArgs e)
        {
       
        }
        private void FillUpARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(FrmFillUpTheARRecord._action)) return;
            if (frmFillUpTheARRecord == null || frmFillUpTheARRecord.IsDisposed)
            {
                frmFillUpTheARRecord = new FrmFillUpTheARRecord();
                frmFillUpTheARRecord.Show();
            }
            frmFillUpTheARRecord.MdiParent = this;
            frmFillUpTheARRecord.Activate();
        }
        /// <summary>
        /// 上传考勤记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UploadARToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(Frm_Upload_AR._action)) return;
            if (frmUploadAR == null || frmUploadAR.IsDisposed)
            {
                frmUploadAR = new Frm_Upload_AR();
                frmUploadAR.Show();
            }
            frmUploadAR.MdiParent = this;
            frmUploadAR.Activate();
        }

        private void TheDaysOfOvertimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
        }

        /// <summary>
        /// 设置加班日
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void setTheDaysOfOverTimeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //查询不可进入。
            if (!ActionHandler.getAction(FrmTheDaysOfOvertime._action)) return;
            if (frmTheDaysOfOvertime == null || frmTheDaysOfOvertime.IsDisposed)
            {
                frmTheDaysOfOvertime = new FrmTheDaysOfOvertime();
                frmTheDaysOfOvertime.Show();
            }
            frmTheDaysOfOvertime.MdiParent = this;
            frmTheDaysOfOvertime.Activate();
        }
        /// <summary>
        /// 设置请假记录
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void SetTheDaysOfAskForLeaveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!ARActionHelper.getAction(FrmAskForLeave._action)) return;
            if (frmAskForLeave == null || frmAskForLeave.IsDisposed)
            {
                frmAskForLeave = new FrmAskForLeave();
                frmAskForLeave.Show();
            }
            frmAskForLeave.MdiParent = this;
            frmAskForLeave.Activate();
        }

        private void ExportARSummaryToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!ARActionHelper.getAction(FrmAnalyzeAR._action)) return;
            if (frmAnalyzeAR == null || frmAnalyzeAR.IsDisposed)
            {
                frmAnalyzeAR = new FrmAnalyzeAR();
                frmAnalyzeAR.Show();
            }
            frmAnalyzeAR.MdiParent = this;
            frmAnalyzeAR.Activate();
        }
    }
}
