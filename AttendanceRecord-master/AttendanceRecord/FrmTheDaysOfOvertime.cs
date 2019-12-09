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
using AttendanceRecord.View;
namespace AttendanceRecord
{
    public partial class FrmTheDaysOfOvertime : Form
    {
        private string defaultDir = Environment.CurrentDirectory + "\\休息日制定\\";
        private FrmRestDay_justConfiguration frmRestDay_justConfiguration = null;
        public static string _action ="admin";
        OracleHelper oH = null;
        private string year_And_Month = string.Empty;
        private string year_Month_Day = string.Empty;
        public FrmTheDaysOfOvertime()
        {
            InitializeComponent();
        }

        private void btnGenerateDefaultRestDays_Click(object sender, EventArgs e)
        {
            if (frmRestDay_justConfiguration == null || frmRestDay_justConfiguration.IsDisposed) {
                frmRestDay_justConfiguration = new FrmRestDay_justConfiguration();
            }
            frmRestDay_justConfiguration.Year_and_month = year_And_Month;
            if (DialogResult.OK.Equals(frmRestDay_justConfiguration.ShowDialog())){
                this.dgv.DataSource = TheDaysOfOvertime.getRestDays(year_And_Month);
                DGVHelper.AutoSizeForDGV(dgv);
            };
        }

        private void mCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            this.year_And_Month = e.Start.ToString("yyyy-MM");
            this.year_Month_Day = e.Start.ToString("yyyy-MM-dd");
            this.dgv.DataSource = TheDaysOfOvertime.getRestDays(year_And_Month);
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void FrmRestDay_Load(object sender, EventArgs e)
        {
            refreshDGV();
        }
        private void refreshDGV() {
            this.dgv.DataSource = TheDaysOfOvertime.getAllRestDays();
            DGVHelper.AutoSizeForDGV(dgv);
        }
        /// <summary>
        /// 删除该行。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void delToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewRow dr = dgv.CurrentRow;
            if (dr == null) return;
            string name = dr.Cells["姓名"].Value.ToString();
            if (string.IsNullOrEmpty(name)) return;
            string rest_day = DateTimeHelper.getDateStr(dr.Cells["休息日"].Value.ToString());
            TheDaysOfOvertime rd = new TheDaysOfOvertime(name, rest_day);
            rd.delRestDay();
            this.dgv.DataSource = TheDaysOfOvertime.getRestDays(year_And_Month);
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.year_Month_Day))
            {
                ShowResult.show(lblResult, "请先选择具体日期", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            TheDaysOfOvertime restDay = new TheDaysOfOvertime("everybody",year_Month_Day);
            restDay.addRestDay();
            this.dgv.DataSource = TheDaysOfOvertime.getRestDays(year_And_Month);
            DGVHelper.AutoSizeForDGV(dgv);
        }
    }
}
