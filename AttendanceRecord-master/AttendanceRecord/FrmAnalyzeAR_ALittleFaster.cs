using AttendanceRecord.Entities;
using AttendanceRecord.Helper;
using AttendanceRecord.View;
using Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;
using AttendanceRecord.GetARSummary;
namespace AttendanceRecord
{
    public partial class FrmAnalyzeAR_ALittleFaster : Form
    {
        public static Queue<int> hwndOfXls_Queue = new Queue<int>();
        public static string _action = "analyzeAR";
        string xlsFilePath = String.Empty;
        private bool _status = true;   //用来表示FrmMain中的noticeIcon是否隐藏。
        string path = string.Empty;
        public FrmAnalyzeAR_ALittleFaster()
        {
            InitializeComponent();
        }
        //用于记录日期值。
        private string YearAndMonthStr = String.Empty;

        /// <summary>
        /// 更新AttendanceRecord中的相关字段。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void mCalendar_DateChanged(object sender, DateRangeEventArgs e)
        {
            YearAndMonthStr = e.Start.ToString("yyyy-MM");
        }

        /// <summary>
        /// 获取考勤记录。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGetARResult_Click(object sender, EventArgs e)
        {
            V_Work_Schedule v_W_S = new V_Work_Schedule();
            v_W_S.WORK_AND_REST_DATE = YearAndMonthStr;
            /*
            if (!v_W_S.ifExistsWS()) {
                return;
            }
            */
            ARResult aRResult = new ARResult(YearAndMonthStr);
            string randomStr = RandomStr.GetRandomString(40, true, true, true, false, "");
            if (aRResult.updateARResult(randomStr) == 0)
            {
                this.dgv.DataSource = MESSAGES.getMSG(randomStr);
                DGVHelper.AutoSizeForDGV(dgv);
                ShowResult.show(lblResult, "异常!", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            ShowResult.show(lblResult, "完成！", true);
            //显示结果
            AttendanceR aR = new AttendanceR();
            this.dgv.DataSource = aR.getARByYearAndMonth(aRResult.Year_And_Month_str);
            DGVHelper.AutoSizeForDGV(dgv);
        }
        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {

        }
        /// <summary>
        /// 导出考勤记录。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportARResult_Click(object sender, EventArgs e)
        {
            //先判断当月的休息日是否指定。
            if (string.IsNullOrEmpty(YearAndMonthStr))
            {
                ShowResult.show(lblResult, "请先设定年月！", false);
                return;
            }
            //判断当前月份有无设定休息日。
            TheDaysOfOvertime restDay = new TheDaysOfOvertime("everybody", YearAndMonthStr);
            if (!restDay.ifHaveTheDayOfOvertime())
            {
                ShowResult.show(lblResult, string.Format(@"请先设定{0}的：休息日！", YearAndMonthStr), false);
                return;
            }
            this.MdiParent.WindowState = FormWindowState.Minimized;
            this.MdiParent.ShowInTaskbar = false;
            ((FrmMainOfAttendanceRecord)this.MdiParent).nfiSystem.Visible = true;
            if (radioBtnSeparate.Checked)
            {
                separateOutputAR();
                //显示完成图标
                timerCompleted.Start();
                return;
            }
            toAWholePiece();
            //显示完成图标。
            timerCompleted.Start();
        }
        private void toAWholePiece()
        {
            killHwndOfXls();
            Queue<int> prefix_Of_Staffs_Queue = V_AR_RESULT.get_Prefix_Staffs(YearAndMonthStr);
            if (prefix_Of_Staffs_Queue.Count == 0)
            {
                ShowResult.show(lblResult, "尚未导入本月的考勤记录！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            MSG msg = new MSG();
            ApplicationClass app = new ApplicationClass();
            //追加Hwnd到队列中.
            int appHwnd = app.Hwnd;
            app.Visible = false;
            Workbook wBook = app.Workbooks.Add(true);
            Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
            string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\考勤汇总";
            int seq = prefix_Of_Staffs_Queue.Dequeue();
            int lastSeq = 0;
            while (prefix_Of_Staffs_Queue.Count >= 1)
                lastSeq = prefix_Of_Staffs_Queue.Dequeue();
            string _fileName = YearAndMonthStr + "_考勤汇总" + seq.ToString() + "-" + lastSeq.ToString() + ".xls";
            if (!xlsFilePath.Contains(":"))
            {
                //导出到Excel中
                xlsFilePath = FileNameDialog.getSaveFileNameWithDefaultDir("考勤汇总：", "*.xls|*.xls", _defaultDir, _fileName);
                if (!xlsFilePath.Contains(@"\"))
                {
                    return;
                }
            }
            //依据前缀和月份获取列表。
            //获取该月该考勤机的出勤人数。
            int AR_Num = V_AttendanceRecord.getARNumByYearAndMonth(YearAndMonthStr);
            if (AR_Num == 0)
            {
                MessageBox.Show("数据源为空，无法导出。", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            int rowMaxCount = AR_Num * 2 + 6;
            int colMaxCount = AttendanceR.get_AR_Days_Num(YearAndMonthStr);
            //写标题
            try
            {
                //每行格式设置，注意标题占一行。
                Range range = wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[rowMaxCount + 1, colMaxCount + 1]);
                //设置单元格为文本。
                range.NumberFormatLocal = "@";
                //水平对齐方式
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //第一行写考勤分析结果。
                wSheet.Cells[1, 1] = YearAndMonthStr + " 考勤分析结果" + seq.ToString() + "-" + lastSeq.ToString();
                //获取该日期详细的考勤记录。
                #region
                V_AttendanceRecord.AR_Properties aR_Properties = V_AttendanceRecord.getARProperties(YearAndMonthStr);
                //第三行：考勤时间
                wSheet.Cells[3, 1] = "考勤时间";
                wSheet.Cells[3, 3] = String.Format(@"{0} ~ {1}",
                                                        aR_Properties.Start_Date,
                                                        aR_Properties.End_Date);
                wSheet.Cells[3, 10] = "制表时间";
                wSheet.Cells[3, 12] = aR_Properties.Tabulation_date;
                #endregion
                List<int> dayList = V_AR_DETAIL.getMonthAL_By_YAndM(YearAndMonthStr);
                //需要统计有多少人。
                int total_num_of_AttendanceR = AttendanceR.get_Total_Num_Of_Staffs_By_YAndM(YearAndMonthStr);
                pb.Value = 0;
                pb.Maximum = dayList.Count + dayList.Count * total_num_of_AttendanceR;
                pb.Visible = true;
                lblPrompt.Text = _fileName + ":";
                lblPrompt.Visible = true;
                //写 此月与考勤相关的日。
                for (int i = 0; i <= dayList.Count - 1; i++)
                {
                    //写该月的具体有哪些日：1，2，3.与考勤相关。
                    wSheet.Cells[4, i + 1] = dayList[i].ToString();
                    pb.Value++;
                }
                //实际出勤天数.
                wSheet.Cells[4, dayList.Count + 1] = "实际出勤天数";
                //事假  
                wSheet.Cells[4, dayList.Count + 2] = "事假";
                //未打卡
                wSheet.Cells[4, dayList.Count + 3] = "未打卡";
                //延点
                wSheet.Cells[4, dayList.Count + 4] = "延点(小时)";
                //迟到
                wSheet.Cells[4, dayList.Count + 5] = "迟到";
                //早退
                wSheet.Cells[4, dayList.Count + 6] = "早退";
                //餐补
                wSheet.Cells[4, dayList.Count + 7] = "餐补";
                string AR_YEAR_AND_Month_Str = String.Empty;
                string AR_Day = string.Empty;
                List<V_AR_DETAIL> v_AR_Detail_Specific_Day_List = null;
                for (int j = 1; j <= dayList.Count; j++)
                {
                    AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                    AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                    v_AR_Detail_Specific_Day_List = V_AR_DETAIL.get_V_AR_Detail_By_Specific_Day(AR_Day);
                    //按日取。
                    for (int i = 0; i <= v_AR_Detail_Specific_Day_List.Count - 1; i++)
                    {
                        V_AR_DETAIL v_AR_Detail = v_AR_Detail_Specific_Day_List[i];
                        if (j == 1)
                        {
                            //第五行写工号。
                            wSheet.Cells[5 + i * 2, 1] = "工号";
                            //获取原始的工号，没有前缀。
                            wSheet.Cells[5 + i * 2, 3] = "'" + v_AR_Detail.Job_number;
                            //9
                            wSheet.Cells[5 + i * 2, 9] = "姓名";
                            //11
                            wSheet.Cells[5 + i * 2, 11] = v_AR_Detail.Name;
                            //19
                            wSheet.Cells[5 + i * 2, 19] = "部门";
                            //21
                            wSheet.Cells[5 + i * 2, 21] = v_AR_Detail.Dept;
                            V_Summary_OF_AR v_summary_of_ar = new V_Summary_OF_AR(v_AR_Detail.Job_number, YearAndMonthStr);
                            System.Data.DataTable dtARSummary = v_summary_of_ar.getSummaryOFAR();
                            //实际出勤天数.
                            wSheet.Cells[6 + i * 2, dayList.Count + 1] = dtARSummary.Rows[0]["AR_DAYS"].ToString();
                            //事假 
                            string vacatioin_total_time = dtARSummary.Rows[0]["VACATION_TOTAL_TIME"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 2] = "0".Equals(vacatioin_total_time) ? "" : vacatioin_total_time;
                            string not_Finger_Print_num = dtARSummary.Rows[0]["NOT_FINGERPRINT_TIMES"].ToString();
                            //未打卡
                            wSheet.Cells[6 + i * 2, dayList.Count + 3] = "0".Equals(not_Finger_Print_num) ? "" : not_Finger_Print_num;
                            string delayTime = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                            //延点
                            wSheet.Cells[6 + i * 2, dayList.Count + 4] = "0.0".Equals(delayTime) ? "" : delayTime;
                            string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                            //迟到
                            wSheet.Cells[6 + i * 2, dayList.Count + 5] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                            string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                            //早退
                            wSheet.Cells[6 + i * 2, dayList.Count + 6] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                            //餐补
                            wSheet.Cells[6 + i * 2, dayList.Count + 7] = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
                        }
                        System.Data.DataTable dt = V_AR_Time_Helper.getARTime(v_AR_Detail.Job_number, AR_Day);
                        string tempStr = String.Empty;
                        int length = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1.ToString().Length;
                        for (int k = 0; k <= dt.Rows.Count - 1; k++)
                        {
                            //先设置颜色.
                            if ("0".Equals(dt.Rows[k]["FLAG"].ToString()))
                            {
                                if ("1".Equals(dt.Rows[k]["COME_LATE_NUM"].ToString())) //迟到
                                {
                                    //先计算单元格已有字符长度。
                                    length = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1.ToString().Length;
                                    //迟到
                                    tempStr = dt.Rows[k]["TIME"].ToString() + (k < dt.Rows.Count - 1 ? "\r\n" : "");
                                    ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 + tempStr;
                                    ((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.Color = -16776961;
                                    continue;
                                }
                                if ("1".Equals(dt.Rows[k]["LEAVE_EARLY_NUM"].ToString()))
                                {
                                    //先计算单元格已有字符长度。
                                    length = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1.ToString().Length;
                                    //早退
                                    tempStr = dt.Rows[k]["TIME"].ToString() + (k < dt.Rows.Count - 1 ? "\r\n" : "");
                                    ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 + tempStr;
                                    //写完即改变前景色。
                                    ((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.Color = -16776961;
                                    continue;
                                }
                                //先计算单元格已有字符长度。
                                length = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1.ToString().Length;
                                //正常
                                //正常上班点.
                                tempStr = dt.Rows[k]["TIME"].ToString() + (k < dt.Rows.Count - 1 ? "\r\n" : "");
                                ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 + tempStr;
                                ((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.Color = System.Drawing.Color.FromArgb(0, 0, 0).ToArgb();
                                continue;
                            }
                            //先计算单元格已有字符长度。
                            length = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1.ToString().Length;
                            //请假点。
                            tempStr = "<" + dt.Rows[k]["TIME"].ToString() + ">" + (k < (dt.Rows.Count - 1) ? "\r\n" : "");
                            ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 = ((Range)wSheet.Cells[6 + i * 2, j]).FormulaR1C1 + tempStr;
                            //((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.Bold = true;
                            //((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.ThemeColor = XlThemeColor.xlThemeColorDark1;
                            //((Range)wSheet.Cells[6 + i * 2, j]).Characters[length + 1, 5].Font.Color = -16776961;
                        }
                        pb.Value++;
                    }
                }
                rowMaxCount = wSheet.UsedRange.Rows.Count;
                //休息日，背景色变为浅绿色。
                for (int j = 1; j <= dayList.Count; j++)
                {
                    bool ifRestDay = false;
                    AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                    AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                    ifRestDay = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                    if (ifRestDay)
                    {
                        //此列背景色改为：
                        /*
                            ange("AF102").Select
                            With Selection.Interior
                                .Pattern = xlSolid
                                .PatternColorIndex = xlAutomatic
                                 .ThemeColor = xlThemeColorAccent3
                                 .TintAndShade = 0.599993896298105
                                .PatternTintAndShade = 0
                            End With
                        End Sub
                        */
                        Range rangeRestDay = wSheet.get_Range(wSheet.Cells[4, j], wSheet.Cells[rowMaxCount, j]);
                        rangeRestDay.Interior.Pattern = XlPattern.xlPatternSolid;
                        rangeRestDay.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                        rangeRestDay.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                        rangeRestDay.Interior.TintAndShade = 0.599993896298105;
                        rangeRestDay.Interior.PatternTintAndShade = 0;
                    }
                }
                //合并第一行
                Range rangeTitle = wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, dayList.Count + 7]);
                rangeTitle.Merge();
                rangeTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rangeTitle.VerticalAlignment = XlVAlign.xlVAlignCenter;
                pb.Visible = false;
                lblPrompt.Visible = false;
                //自动调整列宽
                //range.EntireColumn.AutoFit();
                //设置禁止弹出保存和覆盖的询问提示框
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存excel文档并关闭
                wBook.SaveAs(xlsFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close(true, xlsFilePath, Type.Missing);
                //退出Excel程序
                app.Quit();
                //释放资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                //调用GC的垃圾收集方法
                GC.Collect();
                GC.WaitForPendingFinalizers();
                ShowResult.show(lblResult, "存于: " + xlsFilePath, true);
                timerRestoreTheLblResult.Enabled = true;
                //生成工作安排表。
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示消息:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        private void separateOutputAR()
        {
            string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\考勤汇总";
            killHwndOfXls();
            Queue<int> prefix_Of_Staffs_Queue = V_AR_RESULT.get_Prefix_Staffs(YearAndMonthStr);
            if (prefix_Of_Staffs_Queue.Count == 0)
            {
                ShowResult.show(lblResult, "尚未导入本月的考勤记录！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            string prefix_Of_Staffs = string.Empty;
            //分几个工作表储存。
            while (prefix_Of_Staffs_Queue.Count > 0)
            {
                prefix_Of_Staffs = prefix_Of_Staffs_Queue.Dequeue().ToString();
                string _fileName = YearAndMonthStr + "_考勤汇总" + prefix_Of_Staffs.Substring(0, 1) + ".xls";
                if (!xlsFilePath.Contains(":"))
                {
                    //导出到Excel中
                    xlsFilePath = FileNameDialog.getSaveFileNameWithDefaultDir("考勤汇总：", "*.xls|*.xls", _defaultDir, _fileName);
                    if (!xlsFilePath.Contains(@"\"))
                    {
                        return;
                    }
                }
                else {
                    int index = xlsFilePath.LastIndexOf(string.Format(@"{0}_考勤汇总", YearAndMonthStr));
                    xlsFilePath = xlsFilePath.Remove(index) + _fileName;
                }
                //int index = xlsFilePath.LastIndexOf(string.Format(@"{0}_考勤汇总", YearAndMonthStr));
                //创建文件
                //DirectoryHelper.createFile(xlsFilePath);
                //xlsFilePath = xlsFilePath.Remove(index) + _fileName;
                MyExcel myExcel = new MyExcel(xlsFilePath);
                myExcel.create();
                myExcel.open();
                //追加Hwnd到队列中.
                hwndOfXls_Queue.Enqueue(myExcel.HwndOfApp);
                MSG msg = new MSG();
                //prefix_Of_Staffs = prefix_Of_Staffs_Queue.Dequeue().ToString();
                //依据前缀和月份获取列表。
                //获取第一张表
                Worksheet _firstWS = myExcel.getFirstWorkSheetAfterOpen();
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(_firstWS);
                uEHelper.clearAllContents(_firstWS.UsedRange);
                //清空该文档中的内容。
                Worksheet _secondWS = myExcel.getSecondWorksheetAfterOpen();
                uEHelper = new Usual_Excel_Helper(_secondWS);
                uEHelper.clearAllContents(_secondWS.UsedRange);
                Worksheet _thirdWS = myExcel.getThirdWorksheetAfterOpen();
                uEHelper = new Usual_Excel_Helper(_thirdWS);
                uEHelper.clearAllContents(_thirdWS.UsedRange);
                int machine_no = int.Parse(prefix_Of_Staffs.Substring(0,1));
                //获取一个月内，某考勤机的考勤天数
                int nums_of_ar_days = GetARSummary.GetARSummary.get_nums_of_ar_days(machine_no, YearAndMonthStr);
                //考勤机的考勤天数
                int nums_of_staffs = GetARSummary.GetARSummary.get_nums_of_staffs(machine_no, YearAndMonthStr);
                System.Data.DataTable dt_Staff_Info = GetARSummary.GetARSummary.get_Staff_info(YearAndMonthStr,machine_no);
                System.Data.DataTable dt_AR_Of_Each_Staff = GetARSummary.GetARSummary.get_AR_Of_Each_Staff(YearAndMonthStr, machine_no);
                System.Data.DataTable dt_AR_Summary = GetARSummary.GetARSummary.Get_AR_Summary(YearAndMonthStr, machine_no);
                //隐藏 结果 label;
                lblResult.Visible = false;
                lblPrompt.Visible = true;
                lblPrompt.Text = "考勤机" + prefix_Of_Staffs.Substring(0, 1) + ": 员工信息汇总中...";
                ExcelHelper.saveDtToExcelWithProgressBar(dt_Staff_Info, _firstWS, pb);
                lblPrompt.Text = "考勤机" + prefix_Of_Staffs.Substring(0, 1) + ": 考勤记录汇总中...";
                ExcelHelper.saveDtToExcelWithProgressBar(dt_AR_Of_Each_Staff, _secondWS, pb);
                lblPrompt.Text = "考勤机" + prefix_Of_Staffs.Substring(0, 1) + ": 汇总中...";
                ExcelHelper.saveDtToExcelWithProgressBar(dt_AR_Summary, _thirdWS, pb);
                Microsoft.Office.Interop.Excel.Range range_Src_AR;
                //目标 区域
                Microsoft.Office.Interop.Excel.Range range_desc_AR;
                //AR_Time 在D列存放
                Usual_Excel_Helper uEHelper_firstWS = new Usual_Excel_Helper(_firstWS);
                for (int i = 1; i <= nums_of_staffs; i++)
                {
                    range_Src_AR = ((Microsoft.Office.Interop.Excel.Range)_secondWS.Range[_secondWS.Cells[(i - 1) * (nums_of_ar_days) + 2, 4], _secondWS.Cells[i * nums_of_ar_days + 1]]);
                    range_Src_AR.Copy();
                    //第一张sheet，第4列
                    range_desc_AR = ((Microsoft.Office.Interop.Excel.Range)_firstWS.Cells[i+1,4]);
                    uEHelper_firstWS.pasteByTranspose(range_desc_AR);
                }
                //关闭excel 
                myExcel.save();
                myExcel.close();
            }
            lblResult.Visible = true;
            lblPrompt.Visible = false;
        }
        /// <summary>
        /// 
        /// </summary>
        public static void killHwndOfXls()
        {
            while (hwndOfXls_Queue.Count > 0)
            {
                CmdHelper.killProcessByHwnd(hwndOfXls_Queue.Dequeue());
            }
        }
        private void lblResult_Click(object sender, EventArgs e)
        {
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
            hwndOfXls_Queue.Enqueue(hwnd);
        }
        private void timerShowProgress_Tick(object sender, EventArgs e)
        {

        }
        private void timerCompleted_Tick(object sender, EventArgs e)
        {

        }
        private void FrmAnalyzeAR_ALittleFaster_Load(object sender, EventArgs e)
        {
            if ("AttendanceRecord".Equals(System.Windows.Forms.Application.ProductName))
            {
                path = System.Windows.Forms.Application.StartupPath + "\\";
            }
            else
            {
                path = Environment.CurrentDirectory + "\\" + System.Windows.Forms.Application.ProductVersion + "\\";
            }
        }
    }
}
