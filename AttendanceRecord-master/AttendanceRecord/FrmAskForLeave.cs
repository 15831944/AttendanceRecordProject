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
        DateTime startDateTime;
        DateTime endDateTime;
        int _year;
        int _month;
        int _day;
        int _start_hour;
        int _start_minute;
        int _start_second;

        int _end_hour;
        int _end_minute;
        int _end_second;

        /// <summary>
        /// 依据姓名查询。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbName_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            if (13 == e.KeyChar)
            {
                checkName();
           }
        }

        private bool checkName()
        {
            string name = tbName.Text.Trim();
            if (string.IsNullOrEmpty(name)) {
                tbName.Focus();
                return false;
            } 
            string result = String.Empty;
            if (!"unique".Equals(result =nameHelper.checkName(name))) {
                tbName.Text = "";
                ShowResult.show(lblResult, result, false);
                timerClsResult.Enabled = true;
                tbName.Focus();
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
            if (!checkName()) return;
            if (tbName.Text.Trim() == "") return;
            startDateTime = new DateTime(_year, _month,_day,_start_hour,_start_minute,_start_second);
            endDateTime = new DateTime(_year, _month, _day, _end_hour, _end_minute, _end_second);
            if (startDateTime >= endDateTime) {
                ShowResult.show(lblResult, "结束时间需比起始时间大！", false);
                timerClsResult.Enabled = true;
                return;
            } 
            string startTime = startDateTime.ToString("yyyy-MM-dd HH:mm:ss");
            string endTime = endDateTime.ToString("yyyy-MM-dd HH:mm:ss");
            a_F_L_H = new ASK_For_Leave_Helper(tbName.Text.Trim(), startTime, endTime, tbNO.Text.Trim());
            //先判断是否有日期范围的假条
            if (a_F_L_H.ifExistsAtRange()) {
                ShowResult.show(lblResult, "已存在该日期范围的假条！", false);
                timerClsResult.Enabled = true;
                return;
            }
            if (a_F_L_H.ifExistsVacationAtRange()) {
                ShowResult.show(lblResult, "所设定的范围,涵盖公司休假日,请分段请假！", false);
                timerClsResult.Enabled = true;
                return;
            }
            a_F_L_H.save();
            tbNO.Text = ASK_For_Leave_Helper.getLastedNO();
            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationList();
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void FrmAskForLeave_Load(object sender, EventArgs e)
        {
            tbNO.Text = ASK_For_Leave_Helper.getLastedNO();
            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationList();
            DGVHelper.AutoSizeForDGV(dgv);
            dtPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            timeStartPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,8, 0, 0);
            timeEndPicker.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
        }

        private void timerClsResult_Tick(object sender, EventArgs e)
        {
            lblResult.Text = "";
            lblResult.BackColor = Color.SkyBlue;
            timerClsResult.Enabled = false;
        }

        private void dtStartPicker_ValueChanged(object sender, EventArgs e)
        {
            _year = dtPicker.Value.Year;
            _month = dtPicker.Value.Month;
            _day = dtPicker.Value.Day;
        }

        private void timeStartPicker_ValueChanged(object sender, EventArgs e)
        {
            _start_hour = timeStartPicker.Value.Hour;
            _start_minute = timeStartPicker.Value.Minute;
            if (_start_hour < 8 )
            {
                MessageBox.Show("起始时间点必须从8点开始：", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                timeStartPicker.Value = new DateTime(_year, _month,_day,  8, 0, 0);
                return;
            }
            _start_second = timeStartPicker.Value.Second;
        }

        private void timeEndPicker_ValueChanged(object sender, EventArgs e)
        {
            _end_hour = timeEndPicker.Value.Hour;
            _end_minute = timeEndPicker.Value.Minute;
            if (_end_hour >=17 && _end_minute>0) {
                MessageBox.Show("结束时间最晚为17:00","提示：",MessageBoxButtons.OK,MessageBoxIcon.Information);
                timeEndPicker.Value = new DateTime(_year, _month, _day, 17, 0,0);
                return;
            }
            _end_second = timeEndPicker.Value.Second;
        }
        /// <summary>
        /// 删除该假条
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void delByNOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow dgvR = dgv.CurrentRow;
            if (dgvR == null) return;
            string NO = dgvR.Cells["单号"].Value.ToString();
            ASK_For_Leave_Helper.delTheNO(NO);
            this.dgv.DataSource = ASK_For_Leave_Helper.getAllVacationList();
            DGVHelper.AutoSizeForDGV(dgv);
        }
    }
}
