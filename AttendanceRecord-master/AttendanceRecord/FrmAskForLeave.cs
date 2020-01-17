using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;
using Excel;
using System.Collections;
using AttendanceRecord.Helper;
namespace AttendanceRecord
{
    public partial class FrmAskForLeave : Form
    {
        public FrmAskForLeave() {
            InitializeComponent();
        }
        private ASK_For_Leave_Helper a_F_L_H = null;
        public static string _action = "Ask_For_Leave";
        int _year;
        int _month;
        int _day;


        string year_and_month_str;
        string year_and_month_day_str;

        string startTimeStr;    //格式：  yyyy-MM-dd HH24:MI  to_date(startTimeStr,'yyyy-MM-dd HH24:MI')
        string endTimeStr;      //格式:   yyyy-MM-dd HH24:MI
        /// <summary>
        /// 依据姓名查询。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cbName_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            if (13 == e.KeyChar)
            {
                checkName();
           }
        }

        private bool checkName()
        {
            string name = cbName.Text.Trim();
            if (string.IsNullOrEmpty(name)) {
                cbName.Focus();
                return false;
            } 
            string result = nameHelper.checkName(name);
            if (result.Contains("不")) {
                ShowResult.show(lblResult, result, false);
                timerClsResult.Enabled = true;
                cbName.Focus();
                return false;
            }
            return true;
        }

        /// <summary>
        /// 提交请假
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (!checkName()) { 
                MessageBox.Show(cbName.Text.Trim() + ": 没有在近5个月以内的考勤系统人员名单中。请用备用方案请假!","提示！",MessageBoxButtons.OK,MessageBoxIcon.Information);
                return;
            } 
            if (cbName.Text.Trim() == "") return;
            if (cbTimeSection.Text.Trim() == "") return;
            //startDateTime = new DateTime(_year, _month,_day,_start_hour,_start_minute,_start_second);
            //endDateTime = new DateTime(_year, _month, _day, _end_hour, _end_minute, _end_second);
            /*if (startDateTime >= endDateTime) {
                ShowResult.show(lblResult, "结束时间需比起始时间大！", false);
                timerClsResult.Enabled = true;
                return;
            } 
            */
            //string startTime = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
            //string endTime = endDateTime.ToString("yyyy-MM-dd HH:mm:ss");
             startTimeStr = string.Format(@"{0}-{1}-{2} {3}",
                                                dtPicker.Value.Year.ToString(),
                                                dtPicker.Value.Month.ToString(),
                                                dtPicker.Value.Day.ToString(),
                                                cbTimeSection.Text.ToString().Split('-')[0].Trim().ToString());
             endTimeStr = string.Format(@"{0}-{1}-{2} {3}",
                                                dtPicker.Value.Year.ToString(),
                                                dtPicker.Value.Month.ToString(),
                                                dtPicker.Value.Day.ToString(),
                                                cbTimeSection.Text.ToString().Split('-')[1].Trim().ToString());
            a_F_L_H = new ASK_For_Leave_Helper(cbName.Text.Trim(), startTimeStr, endTimeStr,"");
            //先判断是否有日期范围的假条
            if (a_F_L_H.ifExistsAtRange()) {
                ShowResult.show(lblResult, "已存在该日期范围的假条！", false);
                timerClsResult.Enabled = true;
                return;
            }
            //判断是否设定了加班日
            if (!ifConfigRestDay(year_and_month_str)) {
                ShowResult.show(lblResult, "请先设定加班日！", false);
                timerClsResult.Enabled = true;
                return;
            }
            if (ifTheRestDay(year_and_month_day_str)) {
                ShowResult.show(lblResult, year_and_month_day_str + " :为休息日！", false);
                timerClsResult.Enabled = true;
                return;
            }
            //a_F_L_H.save();
            //tbNO.Text = ASK_For_Leave_Helper.getLastedNO();
            string sqlStr = string.Format(@"insert into Ask_For_Leave(
                                              job_number,
                                                name,
                                                leave_date,
                                              leave_start_time, 
                                              leave_end_time
                                      ) values(
                                        '{0}',
                                        '{1}',
                                        to_date('{2}','yyyy-MM-dd'),
                                        to_date('{3}','yyyy-MM-dd HH24:MI'),
                                        to_date('{4}','yyyy-MM-dd HH24:MI')
                                      )",
                                      cbName.SelectedValue.ToString(),
                                      cbName.Text.Trim(),
                                      year_and_month_day_str,
                                      startTimeStr,
                                      endTimeStr
                                      ); ; 
            OracleDaoHelper.executeSQL(sqlStr);
            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationListByNameAndDate(cbName.Text.Trim());
            DGVHelper.AutoSizeForDGV(dgv);
            cbTimeSection.SelectedIndex = -1;
        }
        private void FrmAskForLeave_Load(object sender, EventArgs e)
        {
            this.cbName.SelectedIndexChanged -= new System.EventHandler(this.cbName_SelectedIndexChanged);
            //tbNO.Text = ASK_For_Leave_Helper.getLastedNO();
           
            
            dtPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            //timeStartPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,8, 0, 0);
            //timeEndPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
            string sqlStr = string.Format(@"select distinct name,job_number
                                              from attendance_record_final
                                              where trunc(finger_print_date,'MM')>=trunc(add_months(sysdate,-3),'MM') 
                                              ORDER BY NLSSORT(name,'NLS_SORT= SCHINESE_PINYIN_M') ASC");
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);

            this.cbName.DataSource = dt;
            cbName.DisplayMember = "name";
            cbName.ValueMember = "job_number";
            this.cbName.SelectedIndexChanged += new System.EventHandler(this.cbName_SelectedIndexChanged);
            cbName.SelectedIndex = -1;

            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationListLastThreeMonths();
            DGVHelper.AutoSizeForDGV(dgv);

        }
        private void timerClsResult_Tick(object sender, EventArgs e)
        {
            lblResult.Text = "";
            lblResult.BackColor = Color.SkyBlue;
            timerClsResult.Enabled = false;
        }
        private void dtPicker_ValueChanged(object sender, EventArgs e)
        {
            _year = dtPicker.Value.Year;
            _month = dtPicker.Value.Month;
            _day = dtPicker.Value.Day;

            year_and_month_str = _year.ToString() + "-" + _month.ToString();
            year_and_month_day_str = _year.ToString() + "-" + _month.ToString() + "-" + _day.ToString();

            loadARDetail();
            //loadAskForLeaveData();
        }
        /// <summary>
        /// 加载本月考勤记录
        /// </summary>
        private void loadARDetail()
        {
            string sqlStr = string.Format(@"select start_date  ""起始日期"",
                                                 end_date ""终止日期"",
                                                 tabulation_time ""制表日期"",
                                                 finger_print_date ""按指纹日期"",
                                                 job_Number  ""工号"",
                                                 name  ""姓名"",
                                                 dept  ""部门"",
                                                 to_char(finger_print_time, 'yy-MM-dd HH24:MI') ""按指纹时间""
                                          from Attendance_Record_Final
                                          where name = '{0}'
                                          and trunc(finger_print_date, 'MM') = to_date('{1}', 'yyyy-MM')
                                          order by finger_print_date asc,
                                                   finger_print_time asc",
                                                   cbName.Text.Trim(),
                                                   year_and_month_str);
            this.dgvARDetail.DataSource = OracleDaoHelper.getDTBySql(sqlStr);
            DGVHelper.AutoSizeForDGV(dgvARDetail);
        }
        /*
        private void timeStartPicker_ValueChanged(object sender, EventArgs e)
        {
            // _start_hour = timeStartPicker.Value.Hour;
            //_start_minute = timeStartPicker.Value.Minute;
            if (_start_hour < 8 )
            {
                MessageBox.Show("起始时间点必须从8点开始：", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                //timeStartPicker.Value = new DateTime(_year, _month,_day,  8, 0, 0);
                return;
            }
            //_start_second = timeStartPicker.Value.Second;
        }
        private void timeEndPicker_ValueChanged(object sender, EventArgs e)
        {
            //_end_hour = timeEndPicker.Value.Hour;
            //_end_minute = timeEndPicker.Value.Minute;
            if (_end_hour >=17 && _end_minute>0) {
                MessageBox.Show("结束时间最晚为17:00","提示：",MessageBoxButtons.OK,MessageBoxIcon.Information);
                //timeEndPicker.Value = new DateTime(_year, _month, _day, 17, 0,0);
                return;
            }
            //_end_second = timeEndPicker.Value.Second;
        }
        */
        /// <summary>
        /// 删除该假条
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void delByNOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow dgvR = dgv.CurrentRow;
            if (dgvR == null) return;
            //string NO = dgvR.Cells["单号"].Value.ToString();
            string name = dgvR.Cells["姓名"].Value.ToString();
            string year_month_day_str = dgvR.Cells["请假日期"].Value.ToString().Substring(0, 10);
            //
            ASK_For_Leave_Helper.delByNameAndMonth(name, year_month_day_str);
            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationListByNameAndDate(cbName.Text.Trim());
            DGVHelper.AutoSizeForDGV(dgv);
        }
        private void cbName_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadAskForLeaveData();
        }
        private void loadAskForLeaveData()
        {
            string sqlStr = string.Format(@"select job_number as ""工号"",
                                             name as ""姓名"",
                                             to_char(leave_date,'yyyy-MM-dd') ""请假日期"",
                                             to_char(leave_start_time,'yyyy-MM-dd HH24:MI')  ""起始时间"",
                                             to_char(leave_end_time,'yyyy-MM-dd HH24:MI') ""终止时间"" 
                                      from Ask_For_Leave
                                      where name = '{0}'
                                        order by leave_date asc", cbName.Text.Trim());   
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            this.dgv.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv);
        }
        #region 判断是否已经设定了休息日。
        private bool ifConfigRestDay(string year_and_month_str)
        {
            string sqlStr = string.Format(@"SELECT 1 
                                            FROM Rest_Day
                                            WHERE trunc(rest_day,'MM') = to_date('{0}','yyyy-MM')", year_and_month_str);
            DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt.Rows.Count > 0 ? true : false;
        }
        #endregion
        private bool ifTheRestDay(string year_and_month_day_str) {
            string sqlStr = string.Format(@"SELECT 1 
                                                FROM REST_DAY
                                                WHERE TRUNC(rest_day,'dd') = to_date('{0}','yyyy-MM-dd')", year_and_month_day_str);
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
        }

        private void cbName_SelectedValueChanged(object sender, EventArgs e)
        {

        }
    }
}
