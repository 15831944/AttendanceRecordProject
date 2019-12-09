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
    public partial class FrmRestDay_Backup : Form
    {
        private string randomStr;
        private string xlsFilePath;
        private string defaultDir = Environment.CurrentDirectory + "\\休息日制定\\";
        public static string _action ="admin";

        private string year_And_Month = string.Empty;

        public FrmRestDay_Backup()
        {
            InitializeComponent();
        }

        private void btnImportRestDay_Click(object sender, EventArgs e)
        {
            randomStr = RandomStr.GetRandomString(33, true, true, true, false, "");
            xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir);
            tbRestDayPath.Text = xlsFilePath;
            //***********************************************************
            List<TheDaysOfOvertime> restDayList = new List<TheDaysOfOvertime>();

            ApplicationClass app = new ApplicationClass();
            app.Visible = false;
            Workbook wBook = null;
            //用于确定本月最后一天.
            //行最大值.
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            try
            {
                wBook = app.Workbooks.Open(xlsFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                if (wBook != null)
                {
                    wBook.Close(true, xlsFilePath, Type.Missing);
                }
                //退出excel
                app.Quit();
                //释放资源
                if (wBook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                //调用GC的垃圾回收方法
                GC.Collect();
                GC.WaitForPendingFinalizers();
                return ;
            }
            Worksheet ws = (Worksheet)(wBook.Worksheets.Item[1]);
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;
            //判断首行是否为 考勤记录表;以此判断此表是否为考勤记录表.
            //检查日期列的值是否符合规范。
            string temp = string.Empty;
            string[] tempArray = { };
            for (int rowIndex = 2; rowIndex <= rowsMaxCount; rowIndex++) {
                temp = ((Range)ws.Cells[rowIndex, 2]).Text.ToString().Trim();
                DateTime dt ;
                if (!DateTime.TryParse(temp, out dt)) {
                    ShowResult.show(lblResult, temp + ": 非时间格式！", false);
                    timerRestoreTheLblResult.Enabled = true;
                    return;
                }
                //先判断是否含有"/"或者"-"
                if (!temp.Contains("/") && !temp.Contains("-"))
                {
                    ShowResult.show(lblResult, "此内容非时间格式: yyyy/MM/dd 或 yyyy-MM-dd！", false);
                    timerRestoreTheLblResult.Enabled = true;
                    return;
                }
                
                tempArray = temp.Split(new char[] { '/', '-' });
                string year = tempArray[0];
                if (!CheckString.checkYear(year)) {
                    ShowResult.show(lblResult, "前４位非年份！", false);
                    timerRestoreTheLblResult.Enabled = true;
                    return;
                }
                string month = tempArray[1];
                if (!CheckString.checkMonth(month)) {
                    ShowResult.show(lblResult, "第６,7位非月份！", false);
                    timerRestoreTheLblResult.Enabled = true;
                    return;
                }
                string day = tempArray[2];
                if (!CheckString.checkDay(day)) {
                    ShowResult.show(lblResult, "第9,10位非天数！", false);
                    timerRestoreTheLblResult.Enabled = true;
                    return;
                }
            }
            string dateStr = string.Empty;
            string name = string.Empty;
            
            //精诚所至，金石为开。
            for (int rowIndex = 2; rowIndex <= rowsMaxCount; rowIndex++)
            {
                dateStr = ((Range)ws.Cells[rowIndex, 2]).Text.ToString().Trim();
                name = ((Range)ws.Cells[rowIndex, 1]).Text.ToString().Trim();
                TheDaysOfOvertime restDay = new TheDaysOfOvertime(name, dateStr);
                restDay.addRestDay();
            }
            //释放对象
            int hwndOfApp = app.Hwnd;
            Tools.CmdHelper.killProcessByHwnd(hwndOfApp);
            //****************************************************************
            tempArray = dateStr.Split(new char[] { '/', '-' });
            this.dgv.DataSource = TheDaysOfOvertime.getRestDays(tempArray[0] + "-" + tempArray[1]);
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void btnGenerateDefaultRestDays_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(this.year_And_Month)) {
                ShowResult.show(lblResult, "请先选择月份！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            OracleParameter param_Name = new OracleParameter("v_Name", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_YearAndMonth = new OracleParameter("v_yearAndMonth", OracleDbType.Varchar2, ParameterDirection.Input);

            OracleHelper oH = OracleHelper.getBaseDao();
            string procedureName = "Generate_Rest_Day";
            param_Name.Value = "everybody";
            param_YearAndMonth.Value = this.year_And_Month;
            OracleParameter[] parameters = new OracleParameter[2] { param_Name, param_YearAndMonth };
            int j = oH.ExecuteNonQuery(procedureName, parameters);
            if (j == 0)
            {
                ShowResult.show(lblResult, "Generate_Rest_Day：生成Rest_Day失败！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            this.dgv.DataSource = TheDaysOfOvertime.getRestDays(year_And_Month);
            DGVHelper.AutoSizeForDGV(dgv);
        }

        private void mCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            this.year_And_Month = e.Start.ToString("yyyy-MM");
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
            string name = dr.Cells["姓名"].ToString();
            string rest_day = dr.Cells["休息日"].ToString();
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
    }
}
