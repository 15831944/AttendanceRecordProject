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
    public partial class FrmFillUpTheARRecord : Form
    {
        private string day = string.Empty;
        private string time = string.Empty;
        public static string _action = "Query";
        string xlsFilePath = String.Empty;
        V_FillUp v_fillUp = null;
        public FrmFillUpTheARRecord()
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
            timerRestoreTheLblResult.Stop();
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
        }

        private void cbName_TextChanged(object sender, EventArgs e)
        {
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

            day = dayPicker.Value.ToString("yyyy-MM-dd");
            time = timerPicker.Value.ToString("hh:mm:ss");

            v_fillUp = new V_FillUp(name, day, time);



            System.Data.DataTable dt = v_fillUp.getARRecordToFillUpByName();
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

        private void cbName_SelectedValueChanged(object sender, EventArgs e)
        {
            btnQuery_Click(sender, e);
        }

        private void FrmFillUpTheARRecord_Load(object sender, EventArgs e)
        {
            this.dayPicker.Value = this.dayPicker.Value.AddMonths(-1);
        }
        /// <summary>
        /// 补卡
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFillUpRecord_Click(object sender, EventArgs e)
        {
            v_fillUp.Day = day;
            v_fillUp.Time = time;

            int fillUpRecordTimes = v_fillUp.getFillUpRecordTimes();

            if (fillUpRecordTimes >= 3) {

                ShowResult.show(lblResult, v_fillUp.Day.Substring(0, 7) + "   " + v_fillUp.Name + "   已经补卡3次了。" , false);
                timerRestoreTheLblResult.Start();
                return;
            }


            string temp = string.Empty;
            if (v_fillUp.ifNotHaveRecordOfMorning() && !v_fillUp.ifNotHaveRecordOfAfternoon()) {
                temp = "上班卡";
            }else if (!v_fillUp.ifNotHaveRecordOfMorning() && v_fillUp.ifNotHaveRecordOfAfternoon()) {
                temp = "下班卡";
            }
            if (!v_fillUp.updateTheRecord()) {
                return;
            }
           
            this.dgv.DataSource = v_fillUp.getARRecordToFillUpByName();
            DGVHelper.AutoSizeForDGV(dgv);
            ShowResult.show(lblResult, day + ": " + temp + "　已补！", true);
            timerRestoreTheLblResult.Start();
       }

        private void dayPicker_ValueChanged(object sender, EventArgs e)
        {
            day = dayPicker.Value.ToString("yyyy-MM-dd");
        }

        private void timerPicker_ValueChanged(object sender, EventArgs e)
        {
            time = timerPicker.Value.ToString("hh:mm:ss");
        }
    }
}
