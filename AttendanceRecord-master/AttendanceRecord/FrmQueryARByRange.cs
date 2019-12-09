using System;
using System.Data;
using System.Windows.Forms;
using Tools;
using System.Collections.Generic;
using AttendanceRecord.Entities;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using AttendanceRecord.Helper;
using Oracle.DataAccess.Client;
using Excel;
using System.Data;
namespace AttendanceRecord
{
    public partial class FrmQueryARByRange : Form
    {
        public static string _action = "Query";
        string xlsFilePath = String.Empty;
        public FrmQueryARByRange()
        {
            InitializeComponent();
        }
        //用于记录日期值。
        private string startDateStr = string.Empty;
        private string endDateStr = string.Empty;

        /// <summary>
        /// 生成工作安排表。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            startDateStr = e.Start.ToString("yyyy-MM-dd");
            endDateStr = e.End.ToString("yyyy-MM-dd");
            AttendanceR aR = new AttendanceR();
            //获取相应的工作安排记录
            System.Data.DataTable dt = aR.getARByRange(startDateStr, endDateStr);
            this.dgv.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv);
        }
        private void mCalendar_DateSelected(object sender, DateRangeEventArgs e)
        {

        }

        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
        }

        private void cbName_TextChanged(object sender, EventArgs e)
        {

            //有问题,这个地方.
            //获取姓名。
            string name = cbName.Text.Trim();
            if (name.Length == 0) return;
            if (name.Length > 1) return;
            string strFirstIndexOfName = name.Substring(0, 1);
            string sqlStr = string.Format(@"select distinct ar.name as name
                                                from attendance_record ar
                                                where ar.name like '{0}%'
                                                and trunc(ar.fingerprint_date,'MM') >= ADD_MONTHS(trunc(sysdate,'MM'),-2)",
                                                strFirstIndexOfName);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            cbName.DataSource = null;
            this.cbName.DataSource = dt;
            cbName.DisplayMember = "name";
            cbName.ValueMember = "name";
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {
            //获取姓名。
            string name = cbName.Text.Trim();
            if (name.Length == 0) return;
            string sqlStr = string.Format(@"select start_date AS ""起始时间"", 
                                                    end_date AS ""终止时间"", 
                                                    tabulation_time AS ""制表时间"", 
                                                    fingerprint_date AS ""按指纹日期"", 
                                                    job_number AS ""工号"", 
                                                    name AS ""姓名"", 
                                                    dept AS ""部门"", 
                                                    to_char(fpt_first_time,'HH24:MI') AS ""上班时间"", 
                                                    to_char(fpt_last_time,'HH24:MI') AS ""下班时间"", 
                                                    come_num AS ""出勤天数"",
                                                    not_finger_print AS ""未打卡时间"",
                                                    ask_for_leave_days AS ""请假天数"", 
                                                    ask_for_leave_type AS ""请假类型"", 
                                                    (case come_late_num
                                                        when 0 THEN N'' 
                                                        else cast(come_late_num as nchar)
                                                        end) AS ""迟到次数"", 
                                                    ( case leave_early_num 
                                                        when 0 THEN N''
                                                        else cast(leave_early_num as nchar)
                                                        end) AS ""早退次数"",  
                                                    delay_time AS ""延时时间"", 
                                                    meal_subsidy AS ""餐补"", 
                                                    random_str AS ""随即字符串"", 
                                                    record_time AS ""记录时间"",
                                                    sheet_name AS ""工作表""
                                                from attendance_record ar
                                                where ar.name ='{0}'
                                                and trunc(ar.fingerprint_date,'DD') between to_date('{1}','YYYY-MM-DD')
                                                    and to_date('{2}','YYYY-MM-DD')
                                                order by ar.fingerprint_date desc",
                                                name,
                                                dtStartDate.Value.ToString("yyyy-MM-dd"),
                                                dtEndDate.Value.ToString("yyyy-MM-dd"));
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            this.dgv.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv);
        }
        /// <summary>
        /// 导出策略
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dgv.Rows.Count == 1) return;
            string dir = Environment.CurrentDirectory + "\\个人考勤记录";
            DirectoryHelper.createDirecotry(dir);
            string xlsFilePath = dir + "\\" +RandomStr.getTimeStamp() + ".xls";
            ExcelHelper.saveDtToExcel((System.Data.DataTable)dgv.DataSource, xlsFilePath);
            ShowResult.show(lblResult, "记录存于" + dir, true);
            timerRestoreTheLblResult.Enabled = true;
        }
    }
}
