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
namespace AttendanceRecord
{
    public partial class FrmGenerateARSummary : Form
    {
        public static string _action = "Generate";
        string xlsFilePath = String.Empty;
        public FrmGenerateARSummary()
        {
            InitializeComponent();
        }
        //用于记录日期值。
        private string YearAndMonthStr = String.Empty;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerateWorkSchedule_Click(object sender, EventArgs e)
        {
            //获取MonthCalendar的值
            //依据Work Summary 生成一张工作安排表
            V_Work_Schedule v_W_S = new V_Work_Schedule();
            V_Work_Schedule._YearAndMonthStr = this.YearAndMonthStr;
            if (!WorkSummary.ifExistsWorkSummary(this.YearAndMonthStr)) {
                ShowResult.show(lblResult, "请先导入5月份考勤记录！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            //生成工作表。
            v_W_S.GenWorkSchedule();
            MSG msg = v_W_S.genExcel(out xlsFilePath);

            ShowResult.show(lblResult, msg.Msg, msg.Flag);
            this.timerRestoreTheLblResult.Enabled = true;
        }
        /// <summary>
        /// 生成工作安排表。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            YearAndMonthStr = e.Start.ToString("yyyy-MM");
            WorkSummary WS = new WorkSummary(YearAndMonthStr);
            V_Work_Schedule._YearAndMonthStr = YearAndMonthStr;
            //先生成工作计划安排
            int affectedCount = WS.generateARSummary();
            //获取相应的工作安排记录
            System.Data.DataTable dt = WS.getARSummary();
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

        private void lblResult_Click(object sender, EventArgs e)
        {
            if (!xlsFilePath.Contains(@"\")) return;
            int hwnd = 0;
            try
            {
                hwnd = ExcelHelper.openBook(xlsFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (hwnd == 0) return;
            V_Work_Schedule.hwndOfXls_Queue.Enqueue(hwnd);
        }
        /// <summary>
        /// 导入工作安排表。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImportWorkSchedule_Click(object sender, EventArgs e)
        {
            string xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择要导入的工作安排表", "*.xls|*.xls", string.Format(@"{0}\{1}", Environment.CurrentDirectory, "工作安排表"));
            if (!xlsFilePath.Contains(@"\")) {
                return;
            }
            MSG msg = new MSG();
            //导入数据的行数.
            int affectedCount = 0;
            ApplicationClass app = new ApplicationClass();
            app.Visible = false;
            Workbook wBook = null;
            //用于确定本月最后一天.
            string day = String.Empty;
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
                msg.Flag = false;
                msg.Msg = ex.ToString();
                return ;
            }
            Worksheet ws = (Worksheet)(wBook.Worksheets.Item[1]);
            if (!((Range)ws.Cells[1, 1]).Text.ToString().Contains("工作安排")) {
                msg.Msg = "此表非工作安排表";
                msg.Flag = false;
                ShowResult.show(lblResult, msg.Msg, msg.Flag);
                return ;
            }
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;
            string C3Str = ((Range)ws.Cells[3, 3]).Text.ToString().Trim();
            string yearAndMonthStr = C3Str.Split('~')[0].Trim().Substring(0,8);
            List<V_I_W_S> v_I_W_S_List = new List<V_I_W_S>();
            //从第二列开始统计
            for (int colIndex=2;colIndex<= colsMaxCount;colIndex++) {
                day = ((Range)ws.Cells[4, colIndex]).Text.ToString().Trim();
                if (String.IsNullOrEmpty(day)) continue;
                //从第5行开始.
                for (int rowIndex = 5;rowIndex<=rowsMaxCount;rowIndex++) {
                    V_I_W_S v_I_W_S = new V_I_W_S();
                    string dept = ((Range)ws.Cells[rowIndex, 1]).Text.ToString().Trim();
                    if (String.IsNullOrEmpty(dept)) continue;
                    v_I_W_S.Dept = dept;
                    v_I_W_S.Date = yearAndMonthStr + day.PadLeft(2,'0');
                    v_I_W_S.BgColor=((Range)(ws.Cells[rowIndex, colIndex])).Interior.Color;
                    v_I_W_S_List.Add(v_I_W_S);
                }
            }
            for (int j=0;j<= v_I_W_S_List.Count-1;j++) {
                affectedCount +=v_I_W_S_List[j].updateWorkSchedule();
            }
            //释放对象
            int hwndOfApp = app.Hwnd;
            Tools.CmdHelper.killProcessByHwnd(hwndOfApp);
            msg.Flag = true;
            msg.Msg = String.Format(@"导入完成;计{0}条.", affectedCount.ToString());
            ShowResult.show(lblResult, msg.Msg, msg.Flag);
        }
        //生成工作安排表。

    }
}
