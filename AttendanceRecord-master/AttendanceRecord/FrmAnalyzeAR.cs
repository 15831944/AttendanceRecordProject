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
using System.Drawing;
namespace AttendanceRecord
{
    public partial class FrmAnalyzeAR : Form
    {
        public static Queue<int> hwndOfXls_Queue = new Queue<int>();
        public static string _action = "analyzeAR";
        string destFilePath= String.Empty;
        private bool _status = true;   //用来表示FrmMain中的noticeIcon是否隐藏。
        string path = string.Empty;
        public FrmAnalyzeAR()
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
            cb_Attendance_Machine.DataSource = null;
            YearAndMonthStr = e.Start.ToString("yyyy-MM");
            #region 获取该日期下的考勤机编号
            /*string sqlStr = string.Format(@"
                                              select * 
                                              from 
                                              (
                                                select distinct(to_char(substr(job_number,1,1))) attendance_machine_no
                                                from Attendance_Record
                                                where trunc(fingerprint_date,'MM') = to_date('{0}','yyyy-MM')  
                                                union
                                                select  wm_concat(distinct(to_char(substr(job_number,1,1))))        
                                                from Attendance_Record
                                                where trunc(fingerprint_date,'MM') = to_date('{0}','yyyy-MM')  
                                              ) temp
                                              order by  length( attendance_machine_no) desc,
                                               attendance_machine_no asc
                                              ",YearAndMonthStr);*/

            /*
             * string sqlStr = string.Format(@"SELECT SUBSTR(SYS_CONNECT_BY_PATH(AR_Machine_Flag,','),2) AR_Machine_Flag
                                                FROM (
                                                     SELECT DISTINCT SUBSTR(job_number,1,1) AR_Machine_Flag
                                                     FROM Attendance_Record 
                                                     WHERE TRUNC(fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')    
                                                )TEMP
                                                CONNECT BY AR_Machine_Flag > PRIOR AR_Machine_Flag 
                                                ORDER BY LENGTH(AR_Machine_Flag) DESC,AR_Machine_Flag ASC", YearAndMonthStr);
            */
            string sqlStr = string.Format(@"SELECT SUBSTR(SYS_CONNECT_BY_PATH(AR_Machine_Flag,','),2) AR_Machine_Flag
                                                FROM (
                                                     SELECT DISTINCT SUBSTR(job_number,1,1) AR_Machine_Flag
                                                     FROM Attendance_Record_Detail 
                                                     WHERE TRUNC(finger_print_date,'MM') = TO_DATE('{0}','yyyy-MM')    
                                                )TEMP
                                                CONNECT BY AR_Machine_Flag > PRIOR AR_Machine_Flag 
                                                ORDER BY LENGTH(AR_Machine_Flag) DESC,AR_Machine_Flag ASC", YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            this.cb_Attendance_Machine.SelectedIndexChanged -=this.cb_Attendance_Machine_SelectedIndexChanged;
            this.cb_Attendance_Machine.DataSource = dt;
            cb_Attendance_Machine.ValueMember = "AR_Machine_Flag";
            this.cb_Attendance_Machine.SelectedIndexChanged += this.cb_Attendance_Machine_SelectedIndexChanged;
            cb_Attendance_Machine.DisplayMember = "AR_Machine_Flag";
            #endregion
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
            if (aRResult.updateARResult(randomStr) == 0) {
                this.dgv.DataSource = MESSAGES.getMSG(randomStr);
                DGVHelper.AutoSizeForDGV(dgv);
                ShowResult.show(lblResult,"异常!",false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            ShowResult.show(lblResult, "完成！", true);
            //显示结果
            AttendanceR aR = new AttendanceR();
            this.dgv.DataSource=aR.getARByYearAndMonth(aRResult.Year_And_Month_str);
            DGVHelper.AutoSizeForDGV(dgv);
        }
        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
        }
        /// <summary>
        /// 导出考勤记录。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportARResult_Click(object sender, EventArgs e)
        {
            this.lblResult.Text = "";
            this.lblResult.BackColor = Color.SkyBlue;
            this.lblResult.Visible = true;
            this.lblPrompt.Visible = false;
            this.pb.Visible = Visible;

            //判断是否存在Excel进程.
            if (CmdHelper.ifExistsTheProcessByName("EXCEL"))
            {
                FrmPrompt frmPrompt = new FrmPrompt();
                frmPrompt.ShowDialog();
            }
            if (string.IsNullOrEmpty(YearAndMonthStr))
            {
                ShowResult.show(lblResult, "请先设定年月！", false);
                return;
            }
            if (string.IsNullOrEmpty(cb_Attendance_Machine.Text.Trim())) {
                ShowResult.show(lblResult, YearAndMonthStr + ": 尚未导入考勤记录！", false);
                return;
            }
            //先判断当月的休息日是否指定。

            //判断当前月份有无设定休息日。
            TheDaysOfOvertime theDaysOfOverTime = new TheDaysOfOvertime("everybody",YearAndMonthStr);
            if (!theDaysOfOverTime.ifHaveTheDayOfOvertime()) {
                ShowResult.show(lblResult,string.Format(@"请先设定{0}的：加班日！",YearAndMonthStr),false);
                return;
            }
            this.MdiParent.WindowState = FormWindowState.Minimized;
            this.MdiParent.ShowInTaskbar = false;
            ((FrmMainOfAttendanceRecord)this.MdiParent).nfiSystem.Visible = true;
            ((FrmMainOfAttendanceRecord)this.MdiParent).nfiSystem.Icon = Properties.Resources.apps;
            /*
            if (radioBtnSeparate.Checked) {
                separateOutputARByPreparedFile();
                //separateOutputAR();
                //
                #region 输出本月考勤数据中同名的用户的考勤汇总，保存至同一张表格中。
                outputSameNameButDifferentAttendanceMachineFalg(YearAndMonthStr);
                #endregion
                //显示完成图标
                timerCompleted.Start();
                return;
            }
            /*
             此处先生成一个汇总。
             */
            //V_Summary_OF_AR.generateAttendanceRecordBriefly(YearAndMonthStr);
            //toAWholePiece();
            toAWholePiece_Previous();
            //显示完成图标。
            timerCompleted.Start();
        }
        /*
        private void toAWholePiece()
        {
            killHwndOfXls();
            Queue<int> prefix_Of_Staffs_Queue = get_Prefix_Of_Staffs_Queue();
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
            string _fileName = YearAndMonthStr + "_考勤汇总" + seq.ToString().Substring(0,1) + "-" + lastSeq.ToString().Substring(0,1) + "_" + TimeHelper.getCurrentTimeStr() + ".xls";
            if (!destFilePath.Contains(":"))
            {
                //导出到Excel中
                destFilePath = FileNameDialog.getSaveFileNameWithDefaultDir("考勤汇总：", "*.xls|*.xls", _defaultDir, _fileName);
                if (!destFilePath.Contains(@"\"))
                {
                    return;
                }
            }
            string sqlStr = string.Format(@"delete from Attendance_Record_Summary");
            OracleDaoHelper.executeSQL(sqlStr);
            //按名称分组  汇总考勤记录到Attendance_Record_Summary。
            saveARToARSummary(YearAndMonthStr);
           //依据前缀和月份获取列表。
           //获取该月该考勤机的出勤人数。
           int AR_Num = V_AttendanceRecord.getARSummaryNumByYearAndMonth(YearAndMonthStr);
                if (AR_Num == 0)
                {
                    MessageBox.Show("数据源为空，无法导出。", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int rowMaxCount = AR_Num * 2 + 6;
                int colMaxCount = AttendanceR.get_A_R_Summary_Days_Num(YearAndMonthStr);
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
                    wSheet.Cells[1, 1] = YearAndMonthStr + " 出勤结果汇总" + seq.ToString() + "-" + lastSeq.ToString();
                    //获取该日期详细的考勤记录。
                    #region
                    V_AttendanceRecord.AR_Properties aR_Properties = V_AttendanceRecord.getPropertiesFromAttendanceRecordSummary(YearAndMonthStr);
                    //第三行：考勤时间
                    wSheet.Cells[3, 1] = "考勤时间";
                    wSheet.Cells[3, 3] = String.Format(@"{0} ~ {1}",
                                                            aR_Properties.Start_Date,
                                                            aR_Properties.End_Date);
                    wSheet.Cells[3, 10] = "制表时间";
                    wSheet.Cells[3, 12] = aR_Properties.Tabulation_date;
                    #endregion
                    List<int> dayList = V_AR_DETAIL.getDaysOfARSummaryByYearAndMonth(YearAndMonthStr );
                    //需要统计有多少人。
                    int total_num_of_AttendanceR = AttendanceR.get_Total_Num_Of_Staffs_Of_A_R_Summary_By_YAndM(YearAndMonthStr);
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
                    //平日加班
                    wSheet.Cells[4, dayList.Count + 4] = "平日延点";
                    //加班日工作时长
                    wSheet.Cells[4, dayList.Count + 5] = "加班日工作时长";
                    //加班合计
                    wSheet.Cells[4, dayList.Count + 6] = "加班合计(小时)";
                    //迟到
                    wSheet.Cells[4, dayList.Count + 7] = "迟到";
                    //早退
                    wSheet.Cells[4, dayList.Count + 8] = "早退";
                    //餐补
                    wSheet.Cells[4, dayList.Count + 9] = "餐补";
                    string AR_YEAR_AND_Month_Str = String.Empty;
                    string AR_Day = string.Empty;
                    List<V_AR_DETAIL> V_A_R_S_Staff_Base_Info_List = null;
                    for (int j = 1; j <= dayList.Count; j++)
                    {
                        AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                        AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                        V_A_R_S_Staff_Base_Info_List = V_AR_DETAIL.get_V_A_R_Summary_Base_Info_By_Specific_Day(AR_Day);
                        //按日取。
                        for (int i = 0; i <= V_A_R_S_Staff_Base_Info_List.Count - 1; i++)
                        {
                            V_AR_DETAIL v_AR_Detail = V_A_R_S_Staff_Base_Info_List[i];
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
                                //获取考勤 汇总: 出勤天数，
                                System.Data.DataTable dtARSummary = v_summary_of_ar.getSummaryOfAttendance_Record_Summary();
                                //实际出勤天数.
                                wSheet.Cells[6 + i * 2, dayList.Count + 1] = dtARSummary.Rows[0]["AR_DAYS"].ToString();
                                //事假 
                                string vacatioin_total_time = dtARSummary.Rows[0]["VACATION_TOTAL_TIME"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 2] = "0".Equals(vacatioin_total_time) ? "" : vacatioin_total_time;
                                string not_Finger_Print_num = dtARSummary.Rows[0]["NOT_FINGERPRINT_TIMES"].ToString();
                                //未打卡
                                wSheet.Cells[6 + i * 2, dayList.Count + 3] = "0".Equals(not_Finger_Print_num) ? "" : not_Finger_Print_num;
                                //平日延时
                                string DELAY_TIMES_OF_Ordinary_Days_str = dtARSummary.Rows[0]["DELAY_TIMES_OF_Ordinary_Days"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 4] = "0.0".Equals(DELAY_TIMES_OF_Ordinary_Days_str) ? "" : DELAY_TIMES_OF_Ordinary_Days_str;
                                //写加班日工作时长
                                string duration_ov_overtime_days_str = dtARSummary.Rows[0]["Duration_Of_Overtime_Days"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 5] = "0.0".Equals(duration_ov_overtime_days_str) ? "" : duration_ov_overtime_days_str;
                                //写总的加班费用。
                                string delayTotalTimes_Str = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 6] = "0.0".Equals(delayTotalTimes_Str) ? "" : delayTotalTimes_Str;

                                string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                                //迟到
                                wSheet.Cells[6 + i * 2, dayList.Count + 7] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                                string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                                //早退
                                wSheet.Cells[6 + i * 2, dayList.Count + 8] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                                //餐补
                                wSheet.Cells[6 + i * 2, dayList.Count + 9] = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
                            }
                            System.Data.DataTable dt = V_AR_Time_Helper.get_AR_Time(v_AR_Detail.Job_number, AR_Day);
                            Range tempRange = ((Range)(wSheet.Cells[6 + i * 2, j]));
                            wSheet.Cells[6 + i * 2, j] = dt.Rows[0]["fpt_first_time"].ToString() + "\r\n" +  dt.Rows[0]["fpt_last_time"].ToString(); 
                            string timeStr = tempRange.Text.ToString().Trim();
                            if (string.IsNullOrEmpty(timeStr))
                            {
                                pb.Value++;
                                if (Math.Round(((decimal)(pb.Value / pb.Maximum) * 100), 2).ToString().Substring(0,2) == "0.3") {
                                    ;
                                }
                                continue;
                            }
                            //获取第一次刷卡的时间.
                            bool comeLate = "1".Equals(dt.Rows[0]["come_late_num"].ToString())? true :false;
                            bool leaveEarly = "1".Equals(dt.Rows[0]["leave_early_num"].ToString())? true :false;
                            //交给一个方法去判断。
                            if (comeLate)
                            {
                                tempRange.Characters[1, 5].Font.Color = -16776961;
                            }
                                    //交给一个方法去判断。
                            if (leaveEarly)
                            {
                                tempRange.Characters[timeStr.Length - 4, 5].Font.Color = -16776961;
                             }
                            if (Math.Round(((decimal)(pb.Value / pb.Maximum) * 100), 2).ToString().Substring(0, 2) == "0.3")
                            {
                                ;
                            }
                        pb.Value++;
                        }
                   }
                    rowMaxCount = wSheet.UsedRange.Rows.Count;
                    //休息日，背景色变为浅绿色。
                    for (int j = 1; j <= dayList.Count; j++)
                    {
                        bool iftheDaysOfOverTime = false;
                        AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                        AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                        iftheDaysOfOverTime = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                        if (iftheDaysOfOverTime)
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
                            /*
                            Range rangetheDaysOfOverTime = wSheet.get_Range(wSheet.Cells[4, j], wSheet.Cells[rowMaxCount, j]);
                            rangetheDaysOfOverTime.Interior.Pattern = XlPattern.xlPatternSolid;
                            rangetheDaysOfOverTime.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                            rangetheDaysOfOverTime.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                            rangetheDaysOfOverTime.Interior.TintAndShade = 0.599993896298105;
                            rangetheDaysOfOverTime.Interior.PatternTintAndShade = 0;
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
                    wBook.SaveAs(destFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wBook.Close(true, destFilePath, Type.Missing);
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
                    ShowResult.show(lblResult, "存于: " + destFilePath, true);
                    timerRestoreTheLblResult.Enabled = true;
                    //生成工作安排表。
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示消息:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
        }
         */
        /// <summary>
        /// 保存考勤汇总到 Attendance_Record_Summary.
        /// 用于处理异机同名用户。
        /// </summary>
        /// <param name="yearAndMonthStr"></param>
        public static void saveARToARSummary(string yearAndMonthStr)
        {
            string procName = "PKG_TO_Export_From_A_R_Summary.Save_AR_To_Summary_Table";
            OracleParameter param_year_and_month_str = new OracleParameter("v_year_and_month_str", OracleDbType.Varchar2, ParameterDirection.Input);
            param_year_and_month_str.Value = yearAndMonthStr;
            OracleParameter[] parameters = new OracleParameter[1] { param_year_and_month_str};
            OracleHelper oH = OracleHelper.getBaseDao();
            oH.ExecuteNonQuery(procName, parameters);
        }

        /// <summary>
        /// 最早的文件备份一份。
        /// </summary>
        private void toAWholePiece_Previous()
        {
            killHwndOfXls();
            Queue<int> prefix_Of_Staffs_Queue = get_Prefix_Of_Staffs_Queue();
            if (prefix_Of_Staffs_Queue.Count == 0)
            {
                ShowResult.show(lblResult, "尚未导入本月的考勤记录！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            MSG msg = new MSG();
            string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\考勤汇总";
            int seq = prefix_Of_Staffs_Queue.Dequeue();
            int lastSeq = 0;
            while (prefix_Of_Staffs_Queue.Count >= 1)
                lastSeq = prefix_Of_Staffs_Queue.Dequeue();
            string _fileName = string.Empty;
            if (0 == lastSeq) {
                _fileName = string.Format(@"{0}_考勤汇总_{1}_{2}.xls", YearAndMonthStr, seq.ToString(),  TimeHelper.getCurrentTimeStr());
            }
            else {
                _fileName = string.Format(@"{0}_考勤汇总_{1}-{2}_{3}.xls", YearAndMonthStr, seq.ToString(), lastSeq.ToString(), TimeHelper.getCurrentTimeStr());
            }
            
            if (!destFilePath.Contains(":"))
            {
                //导出到Excel中
                destFilePath = FileNameDialog.getSaveFileNameWithDefaultDir("考勤汇总：", "*.xls|*.xls", _defaultDir, _fileName);
                if (!destFilePath.Contains(@"\"))
                {
                    return;
                }
            }
            MyExcel myExcel = new MyExcel(destFilePath);
            myExcel.create();
            myExcel.openWithoutAlerts();
            Worksheet wSheet = myExcel.getFirstWorkSheetAfterOpen();
            //依据前缀和月份获取列表。
            //获取该月该考勤机的出勤人数。
            int AR_Num = V_AttendanceRecord.getARNumByAMFlag_YearAndMonth(cb_Attendance_Machine.Text.Trim(),YearAndMonthStr);
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
                if (lastSeq != 0)
                {
                    wSheet.Cells[1, 1] = YearAndMonthStr + " 出勤结果汇总" + seq.ToString() + "-" + lastSeq.ToString();
                }
                else {
                    wSheet.Cells[1, 1] = YearAndMonthStr + " 出勤结果汇总" + seq.ToString();
                }
                //第一行写考勤分析结果。
                
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
                int total_num_of_AttendanceR = AttendanceR.get_Total_Num_Of_Staffs_By_YAndM_And_AMFlag(cb_Attendance_Machine.Text.Trim(),YearAndMonthStr);
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
                //OVERTIME_OF_WEE_HOURS_OF_WD
                //实际出勤天数.
                wSheet.Cells[4, dayList.Count + 1] = "实际出勤总天数";
                //事假  
                wSheet.Cells[4, dayList.Count + 2] = "事假";
                //未打卡
                wSheet.Cells[4, dayList.Count + 3] = "工作日(少打卡)";     //实际指： 周一至周五
                //平日加班
                wSheet.Cells[4, dayList.Count + 4] = "工作日(正班 延点自17:00起)";   
                wSheet.Cells[4, dayList.Count + 5] = "工作日(凌晨 <5:00起至下班)";
                //加班日工作时长
                wSheet.Cells[4, dayList.Count + 6] = "加班日工作时长";
                //加班合计
                wSheet.Cells[4, dayList.Count + 7] = "加班合计(小时)";
                //迟到
                wSheet.Cells[4, dayList.Count + 8] = "迟到";
                //早退
                wSheet.Cells[4, dayList.Count + 9] = "早退";
                //餐补
                wSheet.Cells[4, dayList.Count + 10] = "餐补";
                string AR_YEAR_AND_Month_Str = string.Empty;
                if (!"".Equals(aR_Properties.Start_Date)) {
                    AR_YEAR_AND_Month_Str = AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 7);
                }
                string AR_Day = string.Empty;
                //先删除汇总表
                clear_AR_SUMMARY_FINAL();
                List<V_AR_DETAIL> v_AR_Base_Info_By_Specific_Month_List = null;
                //获取 某月某考勤机的  部门，工号，姓名。
                /*
                 select distinct dept,job_number,name
                                                from Attendance_Record
                                                where substr(job_number,1,1) in ({0})
                                                and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')
                                                order by job_number asc
                 */
                v_AR_Base_Info_By_Specific_Month_List = V_AR_DETAIL.get_V_AR_Detail_By_Attendance_Machine_Flag_And_Specific_Day(cb_Attendance_Machine.Text.Trim(),AR_YEAR_AND_Month_Str);
                for (int j = 1; j <= dayList.Count; j++)
                {
                    
                    AR_Day = AR_YEAR_AND_Month_Str + "-" + dayList[j - 1].ToString().PadLeft(2, '0');
                    //按日取。
                    for (int i = 0; i <= v_AR_Base_Info_By_Specific_Month_List.Count - 1; i++)
                    {
                        V_AR_DETAIL v_AR_Detail = v_AR_Base_Info_By_Specific_Month_List[i];
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
                            v_summary_of_ar.Name = v_AR_Detail.Name;
                            //PKG_Get_Summary_Of_AR.get_summary_of_ar
                            System.Data.DataTable dtARSummary = v_summary_of_ar.getSummaryOFAR();
                            //实际出勤天数.
                            string actual_ar_days_str = dtARSummary.Rows[0]["AR_DAYS"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 1] = "0".Equals(actual_ar_days_str) ? "" : actual_ar_days_str;
                            //事假 
                            string vacatioin_total_time = dtARSummary.Rows[0]["VACATION_TOTAL_TIME"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 2] = "0".Equals(vacatioin_total_time) ? "" : vacatioin_total_time;
                            string not_Finger_Print_num = dtARSummary.Rows[0]["NOT_FINGERPRINT_TIMES"].ToString();
                            //未打卡
                            wSheet.Cells[6 + i * 2, dayList.Count + 3] = "0".Equals(not_Finger_Print_num) ? "" : not_Finger_Print_num;
                            //平日延时
                            string DELAY_TIMES_OF_Ordinary_Days_str = dtARSummary.Rows[0]["DELAY_TIMES_OF_Ordinary_Days"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 4] = "0".Equals(DELAY_TIMES_OF_Ordinary_Days_str) ? "" : DELAY_TIMES_OF_Ordinary_Days_str;
                            //平日凌晨延时
                            string OVERTIME_OF_WEE_HOURS_OF_WD_str = dtARSummary.Rows[0]["OVERTIME_OF_WEE_HOURS_OF_WD"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 5] = "0".Equals(OVERTIME_OF_WEE_HOURS_OF_WD_str) ? "" : OVERTIME_OF_WEE_HOURS_OF_WD_str;
                            //写加班日工作时长
                            string duration_ov_overtime_days_str = dtARSummary.Rows[0]["Duration_Of_Overtime_Days"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 6] = "0".Equals(duration_ov_overtime_days_str) ? "" : duration_ov_overtime_days_str;
                            //写总的加班费用。
                            string delayTotalTimes_Str = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 7] = "0".Equals(delayTotalTimes_Str) ? "" : delayTotalTimes_Str;

                            string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                            //迟到
                            wSheet.Cells[6 + i * 2, dayList.Count + 8] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                            string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                            //早退
                            wSheet.Cells[6 + i * 2, dayList.Count + 9] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                            //餐补
                            string dinnger_subsidy_num_str = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
                            wSheet.Cells[6 + i * 2, dayList.Count + 10] = "0".Equals(dinnger_subsidy_num_str) ?"": dinnger_subsidy_num_str;
                        }
                        System.Data.DataTable dt = V_AR_Time_Helper.getARTimeByJN(v_AR_Detail.Job_number, AR_Day);

                        string tempStr = dt.Rows[0]["fpt_first_time"].ToString() + "\r\n" + dt.Rows[0]["fpt_last_time"].ToString();
                        wSheet.Cells[6 + i * 2, j] = tempStr;
                        if (!string.IsNullOrEmpty(tempStr)) {
                            //获取第一次刷卡的时间.
                            bool comeLate = "1".Equals(dt.Rows[0]["come_late_num"].ToString()) ? true : false;
                            bool leaveEarly = "1".Equals(dt.Rows[0]["leave_early_num"].ToString()) ? true : false;
                            Range tempRange = ((Range)(wSheet.Cells[6 + i * 2, j]));
                            //交给一个方法去判断。
                            if (comeLate)
                            {
                                tempRange.Characters[1, 5].Font.Color = -16776961;
                            }
                            //交给一个方法去判断。
                            if (leaveEarly)
                            {
                                tempRange.Characters[tempStr.Length - 4, 5].Font.Color = -16776961;
                            }
                        }
                        pb.Value++;
                    }
                }
                rowMaxCount = wSheet.UsedRange.Rows.Count;
                //休息日，背景色变为浅绿色。
                for (int j = 1; j <= dayList.Count; j++)
                {
                    bool iftheDaysOfOverTime = false;
                    AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                    AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                    iftheDaysOfOverTime = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                    if (iftheDaysOfOverTime)
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
                        Range rangetheDaysOfOverTime = wSheet.get_Range(wSheet.Cells[4, j], wSheet.Cells[rowMaxCount, j]);
                        rangetheDaysOfOverTime.Interior.Pattern = XlPattern.xlPatternSolid;
                        rangetheDaysOfOverTime.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                        rangetheDaysOfOverTime.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                        rangetheDaysOfOverTime.Interior.TintAndShade = 0.599993896298105;
                        rangetheDaysOfOverTime.Interior.PatternTintAndShade = 0;
                    }
                }
                Range rangeSummary = wSheet.get_Range(wSheet.Cells[4, dayList.Count + 1], wSheet.Cells[4+ v_AR_Base_Info_By_Specific_Month_List.Count*2, dayList.Count + 9]);
                rangeSummary.EntireColumn.AutoFit();
                //Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(wSheet);
                //uEHelper.replace_str(4, dayList.Count + 1, 4 + v_AR_Base_Info_By_Specific_Month_List.Count * 2, dayList.Count + 9,"0","");
                //合并第一行
                Range rangeTitle = wSheet.get_Range(wSheet.Cells[1, 1], wSheet.Cells[1, dayList.Count + 9]);
                rangeTitle.Merge();
                rangeTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rangeTitle.VerticalAlignment = XlVAlign.xlVAlignCenter;
                #region 写记录到sheet2中。
                Worksheet secondSheet = myExcel.getSecondWorksheetAfterOpen();
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(secondSheet);
               
                //实际出勤天数	事假	未打卡	平日延点	加班日工作时长	加班合计(小时)	迟到	早退	餐补
                secondSheet.Cells[1, 1] = "工号";
                secondSheet.Cells[1, 2] = "姓名";
                secondSheet.Cells[1, 3] = "所属部门";
                secondSheet.Cells[1, 4] = "所属组";
                secondSheet.Cells[1, 5] = "实际出勤总天数";
                secondSheet.Cells[1, 6] = "事假";
                secondSheet.Cells[1, 7] = "工作日(少打卡)";
                secondSheet.Cells[1, 8] = "工作日(正班 延点自17:00起)";
                secondSheet.Cells[1, 9] = "工作日(凌晨 <5:00起至下班)";
                secondSheet.Cells[1, 10] = "加班日工作时长";
                secondSheet.Cells[1, 11] = "加班合计(小时)";
                secondSheet.Cells[1, 12] = "迟到";
                secondSheet.Cells[1, 13] = "早退";
                secondSheet.Cells[1, 14] = "餐补";
                secondSheet.Cells[1, 15] = "备注";
                secondSheet.Cells[2, 15] = "所属 部门_班组,数据来自_MES_制卡系统";
                //获取考勤汇总： 所属部门	所属组	实际出勤天数	未打卡	平日延点	加班日工作时长	加班合计(小时)	迟到	早退	餐补	备注
                //PKG_AR_SUMMARY.getARSummary
                System.Data.DataTable dtARSummaryFinal = ARSummaryFinal.getARSummaryFinal(YearAndMonthStr);
                uEHelper.ChangeFormatToText(uEHelper.getRange("A1","A"+ (dtARSummaryFinal.Rows.Count +1).ToString()));
                int rowIndex = 2;
                for (int i = 0; i <= dtARSummaryFinal.Rows.Count -1; i++)
                {
                    secondSheet.Cells[rowIndex, 1] = dtARSummaryFinal.Rows[i]["job_number"].ToString();
                    string name = dtARSummaryFinal.Rows[i]["name"].ToString();
                    secondSheet.Cells[rowIndex, 2] = name;
                    #region
                    if (Program.flag_open_mesSqlConn) {
                        string sqlStr = string.Format(@"SELECT TOP 1 HR_E.EmployeeName,
				                                        HR_T.TeamName TeamName,
				                                        HR_D.deptName DeptName
		                                        from HR_Employee HR_E INNER JOIN HR_Team HR_T
					                                        ON HR_E.TeamID = HR_T.TeamID
						                                        INNER JOIN HR_Depart HR_D
					                                        ON HR_E.deptID = HR_D.deptID
                                                WHERE HR_E.EmployeeName = '{0}'
                                                    AND NOT (ISNULL(HR_E.TeamName,'') = '' AND ISNULL(HR_T.TeamName,'') = '') 
		                                        ORDER BY HR_E.EmployeeCode ASC", name);
                        System.Data.DataTable dtDeptAndTeamInfo = SqlDaoHelper.getDTBySql(sqlStr);
                        if (dtDeptAndTeamInfo.Rows.Count != 0)
                        {
                            secondSheet.Cells[rowIndex, 3] = dtDeptAndTeamInfo.Rows[0]["DeptName"].ToString();
                            secondSheet.Cells[rowIndex, 4] = dtDeptAndTeamInfo.Rows[0]["TeamName"].ToString();
                        }
                    }
                    #endregion
                    string days_of_real_attendance_str = dtARSummaryFinal.Rows[i]["days_of_real_attendance"].ToString();
                    secondSheet.Cells[rowIndex, 5] = "0".Equals(days_of_real_attendance_str) ? "" : days_of_real_attendance_str;
                    //请假时长
                    string vataction_total_time_str = dtARSummaryFinal.Rows[i]["vacation_total_time"].ToString();
                    secondSheet.Cells[rowIndex, 6] = "0".Equals(vataction_total_time_str) ? "" : vataction_total_time_str;
                    string not_finger_print_str = dtARSummaryFinal.Rows[i]["not_finger_print"].ToString();
                    secondSheet.Cells[rowIndex, 7] = "0".Equals(not_finger_print_str) ? "" : not_finger_print_str;
                    string overtime_of_workday_str = dtARSummaryFinal.Rows[i]["overtime_of_workday"].ToString();
                    secondSheet.Cells[rowIndex, 8] = "0".Equals(overtime_of_workday_str) ? "" : overtime_of_workday_str;
                    string OVERTIME_OF_WEE_HOURS_OF_WD_str = dtARSummaryFinal.Rows[i]["OVERTIME_OF_WEE_HOURS_OF_WD"].ToString();
                    secondSheet.Cells[rowIndex, 9] = "0".Equals(OVERTIME_OF_WEE_HOURS_OF_WD_str) ? "" : OVERTIME_OF_WEE_HOURS_OF_WD_str;
                    string overtime_of_restday_str = dtARSummaryFinal.Rows[i]["overtime_of_restday"].ToString();
                    secondSheet.Cells[rowIndex, 10] = "0".Equals(overtime_of_restday_str) ? "" : overtime_of_restday_str;
                    string total_overtime_str = dtARSummaryFinal.Rows[i]["total_overtime"].ToString();
                    secondSheet.Cells[rowIndex, 11] = "0".Equals(total_overtime_str) ? "" : total_overtime_str;
                    string come_late_str = dtARSummaryFinal.Rows[i]["come_late_num"].ToString();
                    secondSheet.Cells[rowIndex, 12] = "0".Equals(come_late_str) ? "" : come_late_str;
                    string leave_early_num_str = dtARSummaryFinal.Rows[i]["leave_early_num"].ToString();
                    secondSheet.Cells[rowIndex, 13] = "0".Equals(leave_early_num_str) ? "" : leave_early_num_str;
                    string meal_subsidy_str = dtARSummaryFinal.Rows[i]["meal_subsidy"].ToString();
                    secondSheet.Cells[rowIndex, 14] = "0".Equals(meal_subsidy_str) ? "" : meal_subsidy_str;
                    rowIndex++;
                }
                Range secondSheetUsedRange = secondSheet.UsedRange;
                secondSheetUsedRange.EntireColumn.AutoFit();
                secondSheet.Select();
                #endregion
                myExcel.save();
                myExcel.close();
                pb.Visible = false;
                lblPrompt.Visible = false;
                //自动调整列宽
                //range.EntireColumn.AutoFit();
                //设置禁止弹出保存和覆盖的询问提示框

                //调用GC的垃圾收集方法
                GC.Collect();
                GC.WaitForPendingFinalizers();
                ShowResult.show(lblResult, "存于: " + destFilePath, true);
                timerRestoreTheLblResult.Enabled = true;
                timerCompleted.Start();

                //生成工作安排表。
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示消息:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                throw new Exception(ex.ToString());
            }
        }
        /// <summary>
        /// 删除汇总表格表。
        /// </summary>
        private void clear_AR_SUMMARY_FINAL()
        {
            string sqlStr = string.Format(@"delete from AR_Summary_Final");
            OracleDaoHelper.executeSQL(sqlStr);
        }

        private Queue<int> get_Prefix_Of_Staffs_Queue()
        {
            Queue<int> prefix_of_Staffs_Queue = new Queue<int>();
            if(string.IsNullOrEmpty(cb_Attendance_Machine.Text.Trim()))
            return prefix_of_Staffs_Queue;
            string[] machine_no_array = cb_Attendance_Machine.Text.Split(',');
            foreach (string machine_no in machine_no_array) {
                //prefix_of_Staffs_Queue.Enqueue(int.Parse(machine_no.PadLeft(9, machine_no.ToCharArray()[0])));
                prefix_of_Staffs_Queue.Enqueue(int.Parse(machine_no));
            }
            return prefix_of_Staffs_Queue;
        }
        /// <summary>
        /// 
        /// </summary>
        private void separateOutputAR()
        {
            killHwndOfXls();
            Queue<int> prefix_Of_Staffs_Queue = get_Prefix_Of_Staffs_Queue();
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
                killHwndOfXls();
                MSG msg = new MSG();
                ApplicationClass app = new ApplicationClass();
                //追加Hwnd到队列中.
                hwndOfXls_Queue.Enqueue(app.Hwnd);
                app.Visible = false;
                Workbook wBook = app.Workbooks.Add(true);
                Worksheet wSheet = (Worksheet)wBook.Worksheets[1];
                string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\考勤汇总";
                prefix_Of_Staffs = prefix_Of_Staffs_Queue.Dequeue().ToString();
                string _fileName = YearAndMonthStr + "_考勤汇总" + prefix_Of_Staffs.Substring(0, 1) + "_" + TimeHelper.getCurrentTimeStr() + ".xls";
                if (!destFilePath.Contains(":"))
                {
                    //导出到Excel中
                    destFilePath = FileNameDialog.getSaveFileNameWithDefaultDir("考勤汇总：", "*.xls|*.xls", _defaultDir, _fileName);
                    if (!destFilePath.Contains(@"\"))
                    {
                        return;
                    }
                }
                int index = destFilePath.LastIndexOf(string.Format(@"{0}_考勤汇总", YearAndMonthStr));
                destFilePath = destFilePath.Remove(index) + _fileName;
                //依据前缀和月份获取列表。
                //获得与该月该考勤机同名的所有员工。
                int AR_Num = V_AttendanceRecord.getARNum(YearAndMonthStr, prefix_Of_Staffs);
                if (AR_Num == 0)
                {
                    MessageBox.Show("数据源为空，无法导出。", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                //确定行数目。
                int rowMaxCount = AR_Num * 2 + 6;
                //该月该考勤机共有多少个日期。
                int colMaxCount = AttendanceR.get_AR_Days_Num(YearAndMonthStr, prefix_Of_Staffs);
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
                    wSheet.Cells[1, 1] = YearAndMonthStr + " 考勤分析结果" + prefix_Of_Staffs.Substring(0, 1);
                    //获取该日期详细的考勤记录。
                    #region
                    V_AttendanceRecord.AR_Properties aR_Properties = V_AttendanceRecord.getARProperties(YearAndMonthStr, prefix_Of_Staffs);
                    //第三行：考勤时间
                    wSheet.Cells[3, 1] = "考勤时间";
                    wSheet.Cells[3, 3] = String.Format(@"{0} ~ {1}",
                                                            aR_Properties.Start_Date,
                                                            aR_Properties.End_Date);
                    wSheet.Cells[3, 10] = "制表时间";
                    wSheet.Cells[3, 12] = aR_Properties.Tabulation_date;
                    #endregion
                    //获取该月该考勤机记录的考勤日期。

                    List<int> dayList = V_AR_DETAIL.getMonthAL_By_YAndM_And_Prefix_JN(YearAndMonthStr, prefix_Of_Staffs);
                    //与该月该考勤机所记录的同名的员工个数。
                    int total_num_of_AttendanceR = AttendanceR.get_Total_Num_Of_Staffs_By_YAndM_And_PreJN(YearAndMonthStr, prefix_Of_Staffs);
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
                    //平日加班
                    wSheet.Cells[4, dayList.Count + 4] = "平日延点";
                    //加班日工作时长
                    wSheet.Cells[4, dayList.Count + 5] = "加班日工作时长";
                    //加班合计
                    wSheet.Cells[4, dayList.Count + 6] = "加班合计(小时)";
                    //迟到
                    wSheet.Cells[4, dayList.Count + 7] = "迟到";
                    //早退
                    wSheet.Cells[4, dayList.Count + 8] = "早退";
                    //餐补
                    wSheet.Cells[4, dayList.Count + 9] = "餐补";
                    string AR_YEAR_AND_Month_Str = String.Empty;
                    string AR_Day = string.Empty;
                    List<V_AR_DETAIL> v_AR_Detail_Specific_Day_List = null;
                    for (int j = 1; j <= dayList.Count; j++)
                    {
                        AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                        AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                        v_AR_Detail_Specific_Day_List = V_AR_DETAIL.get_V_AR_Detail_By_Specific_Day_And_PreJN(AR_Day, prefix_Of_Staffs);
                        //按日取。
                        for (int i = 0; i <= v_AR_Detail_Specific_Day_List.Count - 1; i++)
                        {
                            V_AR_DETAIL v_AR_Detail = v_AR_Detail_Specific_Day_List[i];
                            if (j == 1)
                            {
                                //第五行写工号。
                                wSheet.Cells[5 + i * 2, 1] = "工号";
                                //获取原始的工号，没有前缀。
                                wSheet.Cells[5 + i * 2, 3] = v_AR_Detail._Previous_Job_number;
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
                                //平日延时
                                string DELAY_TIMES_OF_Ordinary_Days_str = dtARSummary.Rows[0]["DELAY_TIMES_OF_Ordinary_Days"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 4] = "0.0".Equals(DELAY_TIMES_OF_Ordinary_Days_str) ? "" : DELAY_TIMES_OF_Ordinary_Days_str;
                                //写加班日工作时长
                                string duration_ov_overtime_days_str = dtARSummary.Rows[0]["Duration_Of_Overtime_Days"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 5] = "0.0".Equals(duration_ov_overtime_days_str) ? "" : duration_ov_overtime_days_str;
                                //写总的加班费用。
                                string delayTotalTimes_Str = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                                wSheet.Cells[6 + i * 2, dayList.Count + 6] = "0.0".Equals(delayTotalTimes_Str) ? "" : delayTotalTimes_Str;

                                string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                                //迟到
                                wSheet.Cells[6 + i * 2, dayList.Count + 7] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                                string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                                //早退
                                wSheet.Cells[6 + i * 2, dayList.Count + 8] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                                //餐补
                                wSheet.Cells[6 + i * 2, dayList.Count + 9] = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
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
                        bool iftheDaysOfOverTime = false;
                        AR_YEAR_AND_Month_Str = aR_Properties.Start_Date.Substring(0, 8);
                        AR_Day = AR_YEAR_AND_Month_Str + dayList[j - 1].ToString().PadLeft(2, '0');
                        iftheDaysOfOverTime = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                        if (iftheDaysOfOverTime)
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
                            Range rangetheDaysOfOverTime = wSheet.get_Range(wSheet.Cells[4, j], wSheet.Cells[rowMaxCount, j]);
                            rangetheDaysOfOverTime.Interior.Pattern = XlPattern.xlPatternSolid;
                            rangetheDaysOfOverTime.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                            rangetheDaysOfOverTime.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                            rangetheDaysOfOverTime.Interior.TintAndShade = 0.599993896298105;
                            rangetheDaysOfOverTime.Interior.PatternTintAndShade = 0;
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
                    wBook.SaveAs(destFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wBook.Close(true, destFilePath, Type.Missing);
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
                    ShowResult.show(lblResult, "存于: " + destFilePath, true);
                    timerRestoreTheLblResult.Enabled = true;
                    //生成工作安排表。
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "提示消息:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
            }
        }
        public static void killHwndOfXls()
        {
          CmdHelper.killProcessByHwndQueue(hwndOfXls_Queue);
        }
        /// <summary>
        /// 依据准备文件导出。
        /// </summary>
        private void separateOutputARByPreparedFile(){
            Queue<int> prefix_Of_Staffs_Queue = get_Prefix_Of_Staffs_Queue() ;
            if (prefix_Of_Staffs_Queue.Count == 0)
            {
                ShowResult.show(lblResult, "尚未导入本月的考勤记录！", false);
                timerRestoreTheLblResult.Enabled = true;
                return;
            }
            string prefix_Of_Staffs = string.Empty;
            System.Data.DataTable dtdd = Have_A_Rest_Helper.getDaysOfRestDay(YearAndMonthStr);
            //分几个工作表储存。
            while (prefix_Of_Staffs_Queue.Count > 0)
            {
                //打开准备文件。
                prefix_Of_Staffs = prefix_Of_Staffs_Queue.Dequeue().ToString();
                //获取考勤机器标识
                string attendanceMachineFlag = prefix_Of_Staffs.Substring(0, 1);
                //lblPrompt 用于进度条提示。
                lblPrompt.Text = string.Format(@"{0}_{1} 汇总进度：",YearAndMonthStr, attendanceMachineFlag);
                //定义准备文件名。
                string preparedFileName = YearAndMonthStr + "_" + attendanceMachineFlag + ".xls";
                //复制准备文件到考勤汇总中。
                string _preparedFilePath = System.Windows.Forms.Application.StartupPath + "\\prepared\\" + preparedFileName;
                string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\考勤汇总";
                //目标文件名
                string destFileName = YearAndMonthStr + "_" + attendanceMachineFlag + "_" + TimeHelper.getCurrentTimeStr() + ".xls";
                //复制文件到目的目录
                CmdHelper.copyFileToDestDirWithNewFileName(_preparedFilePath, _defaultDir, destFileName);
                //目标文件路径
                destFilePath = _defaultDir + "\\" + destFileName;
                //打开目标文件
                MyExcel myExcel = new MyExcel(destFilePath);
                myExcel.openWithoutAlerts();
                //获取第一个sheet 
                Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
                //去除掉文件中的0
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstSheet);
                int AR_Num = V_AttendanceRecord.getARNum(YearAndMonthStr, prefix_Of_Staffs);
                if (AR_Num == 0)
                {
                    MessageBox.Show("数据源为空，无法导出。", "提示：", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                firstSheet.Cells[1, 1] = YearAndMonthStr + " 考勤分析结果" + prefix_Of_Staffs.Substring(0, 1);
                #region 对加班日  以绿色 进行 标识。
                //汇总结果的终止行，列号。
                int summary_end_rowIndex = firstSheet.UsedRange.Rows.Count;
                List<int> dayList = V_AR_DETAIL.getMonthAL_By_YAndM_And_Prefix_JN(YearAndMonthStr, prefix_Of_Staffs);
                int daysCount = dayList.Count;
                int maxColBefourBlankCellInThe4thRow = uEHelper.getMaxColIndexBeforeBlankCellInSepcificRow(4);
                int maxColIndexOfSrcRange = daysCount;
                //休息日，背景色变为浅绿色。
                for (int colIndex = 1; colIndex <= maxColIndexOfSrcRange; colIndex++)
                {
                    bool iftheDaysOfOverTime = false;
                    string dayStr = ((Range)(firstSheet.Cells[4, colIndex])).Text.ToString();
                    string AR_Day = YearAndMonthStr + dayStr.PadLeft(2, '0');
                    iftheDaysOfOverTime = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                    if (iftheDaysOfOverTime)
                    {
                        Range rangetheDaysOfOverTime = firstSheet.get_Range(firstSheet.Cells[4, colIndex], firstSheet.Cells[summary_end_rowIndex, colIndex]);
                        rangetheDaysOfOverTime.Interior.Pattern = XlPattern.xlPatternSolid;
                        rangetheDaysOfOverTime.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                        rangetheDaysOfOverTime.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                        rangetheDaysOfOverTime.Interior.TintAndShade = 0.599993896298105;
                        rangetheDaysOfOverTime.Interior.PatternTintAndShade = 0;
                    }
                }
                #endregion
                //汇总结果的起始行，列号。
                int summary_start_rowIndex = 4;

                int summary_start_colIndex = maxColBefourBlankCellInThe4thRow > daysCount ? maxColBefourBlankCellInThe4thRow : daysCount;
                summary_start_colIndex++;

                int summary_end_colIndex = summary_start_colIndex + 8;
                //每行格式设置，注意标题占一行。
                Range summaryRange = firstSheet.get_Range(firstSheet.Cells[summary_start_rowIndex, summary_start_colIndex], firstSheet.Cells[summary_end_rowIndex, summary_end_colIndex]);
                //设置单元格为文本。
                summaryRange.NumberFormatLocal = "@";
                //水平对齐方式
                summaryRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                #region  写汇总标题
                //                //第一行写考勤分析结果。
                //实际出勤天数.
                firstSheet.Cells[4, summary_start_colIndex] = "实际出勤天数";
                //事假  
                firstSheet.Cells[4, summary_start_colIndex + 1] = "事假";
                //未打卡
                firstSheet.Cells[4, summary_start_colIndex + 2] = "未打卡";
                //平日加班
                firstSheet.Cells[4, summary_start_colIndex + 3] = "平日延点";
                //加班日工作时长
                firstSheet.Cells[4, summary_start_colIndex + 4] = "加班日工作时长";
                //加班合计
                firstSheet.Cells[4, summary_start_colIndex + 5] = "加班合计(小时)";
                //迟到
                firstSheet.Cells[4, summary_start_colIndex + 6] = "迟到";
                //早退
                firstSheet.Cells[4, summary_start_colIndex + 7] = "早退";
                //餐补
                firstSheet.Cells[4, summary_start_colIndex + 8] = "餐补";
                #endregion
                pb.Maximum = 0;
                pb.Value = 0;
                pb.Visible = true;
                pb.Maximum += summary_end_rowIndex - 5 + 1;
                pb.Maximum += (summary_end_rowIndex - 5 + 1) * maxColIndexOfSrcRange;
                #region 写汇总结果
                for (int rowIndex = 5; rowIndex <= summary_end_rowIndex;rowIndex++)
                {
                    if (1 == rowIndex % 2) {
                        //奇数行,不处理，汇总写在偶数行
                        pb.Value++;
                        continue;
                    }
                    //获取工号。
                    int rowIndexOfJN = rowIndex - 1;
                    string original_jn = ((Range)(firstSheet.Cells[rowIndexOfJN, 3])).Text.ToString().Trim();
                    if (string.IsNullOrEmpty(original_jn))
                    {
                        pb.Value++;
                        continue;
                    }
                    original_jn = prefix_Of_Staffs + original_jn.PadLeft(3, '0');
                    V_Summary_OF_AR v_summary_of_ar = new V_Summary_OF_AR(original_jn, YearAndMonthStr);
                    System.Data.DataTable dtARSummary = v_summary_of_ar.getSummaryOFAR();

                    //实际出勤天数.
                    firstSheet.Cells[rowIndex, summary_start_colIndex] = dtARSummary.Rows[0]["AR_DAYS"].ToString();

                    //事假 
                    string vacatioin_total_time = dtARSummary.Rows[0]["VACATION_TOTAL_TIME"].ToString();
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 1] = "0".Equals(vacatioin_total_time) ? "" : vacatioin_total_time;

                    string not_Finger_Print_num = dtARSummary.Rows[0]["NOT_FINGERPRINT_TIMES"].ToString();
                    //未打卡
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 2] = "0".Equals(not_Finger_Print_num) ? "" : not_Finger_Print_num;
                    //平日延时
                    string DELAY_TIMES_OF_Ordinary_Days_str = dtARSummary.Rows[0]["DELAY_TIMES_OF_Ordinary_Days"].ToString();
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 3] = "0.0".Equals(DELAY_TIMES_OF_Ordinary_Days_str) ? "" : DELAY_TIMES_OF_Ordinary_Days_str;
                    //写加班日工作时长
                    string duration_ov_overtime_days_str = dtARSummary.Rows[0]["Duration_Of_Overtime_Days"].ToString();
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 4] = "0.0".Equals(duration_ov_overtime_days_str) ? "" : duration_ov_overtime_days_str;
                    //写总的加班费用。
                    string delayTotalTimes_Str = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 5] = "0.0".Equals(delayTotalTimes_Str) ? "" : delayTotalTimes_Str;

                    string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                    //迟到
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 6] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                    string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                    //早退
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 7] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                    //餐补
                    firstSheet.Cells[rowIndex, summary_start_colIndex + 8] = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
                    pb.Value++;
                }
                #endregion
                #region 标识 存在  迟到或者早退的单元格。
                for (int rowIndex = 5; rowIndex <= summary_end_rowIndex;rowIndex++)
                {
                    if (1 == rowIndex % 2) {
                        //奇数行：工号行，不做处理。
                        pb.Value += maxColIndexOfSrcRange;
                        continue;
                    }
                    //偶数行。
                    for (int colIndex = 1; colIndex <= maxColIndexOfSrcRange; colIndex++)
                    {
                        //判断是否为加班日，加班日不计算迟到，早退。
                        //在第4行，取dd
                        string ddStr = ((Range)(firstSheet.Cells[4, colIndex])).Text.ToString().Trim();
                        bool isRestDay = false;
                        for (int i=0;i<=dtdd.Rows.Count-1;i++) {
                            if (ddStr.Equals(dtdd.Rows[i]["dd"].ToString())){
                                //若为加班日，不做标识。
                                isRestDay = true;
                                break;
                            }
                        }
                        //是加班日，不计算迟到，早退。
                        if (isRestDay) {
                            pb.Value++;
                            continue;
                        }
                        Range tempRange = ((Range)(firstSheet.Cells[rowIndex, colIndex]));
                        string timeStr = tempRange.Text.ToString().Trim();
                        if (string.IsNullOrEmpty(timeStr))
                        {
                            pb.Value++;
                            continue;
                        }
                        //获取第一次刷卡的时间.
                        bool comeLate = false;
                        bool leaveEarly = false;
                        AttendanceTimeHelper.judgeIfComeLateOrLeaveEarly(timeStr, out comeLate, out leaveEarly);
                        //交给一个方法去判断。
                        if (comeLate)
                        {
                            tempRange.Characters[1, 5].Font.Color = -16776961;
                        }
                        //交给一个方法去判断。
                        if (leaveEarly)
                        {
                            tempRange.Characters[timeStr.Length-4, 5].Font.Color = -16776961;
                        }
                        pb.Value++;
                    }
                }
                #endregion
              
                Range rangToReplace = firstSheet.get_Range(firstSheet.Cells[6, summary_start_colIndex], firstSheet.Cells[summary_end_rowIndex, summary_start_colIndex + 8]);
                uEHelper.replace_str(6, summary_start_colIndex, summary_end_rowIndex, summary_start_colIndex + 8, "0", "");
                //合并第一行
                Range rangeTitle = firstSheet.get_Range(firstSheet.Cells[1, 1], firstSheet.Cells[1, summary_start_colIndex + 8]);
                rangeTitle.Merge();
                rangeTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                rangeTitle.VerticalAlignment = XlVAlign.xlVAlignCenter;
                //对汇总区域进行列宽自适应。
                myExcel.saveAndColumnsAutoFit(summaryRange);
                myExcel.close();
                ShowResult.show(lblResult, "存于: " + destFilePath, true);
                timerRestoreTheLblResult.Enabled = true;
                pb.Visible = false;
            }
        }
        /*
        /// <summary>
        /// 将某月的考勤记录中同名的记录，汇总输出到一个表格中。
        /// </summary>
        /// <param name="yearAndMonthStr"></param>
        /// 
        private void outputSameNameButDifferentAttendanceMachineFalg(string yearAndMonthStr)
        {
            string sqlStr = string.Format(@"SELECT distinct dept,job_number,name
                                                FROM Attendance_Record AR
                                                WHERE TRUNC(Fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                                AND AR.name = ANY
                                                (
                                                  SELECT name 
                                                   FROM  
                                                   (
                                                      SELECT row_number() over(partition by name order by job_number asc) row_id,
                                                             job_number,
                                                             name               
                                                      FROM 
                                                      (
                                                           SELECT DISTINCT name,
                                                                      job_number
                                                           from Attendance_Record AR
                                                           where trunc(fingerprint_date,'MM') = to_date('{0}','yyyy-MM')
                                                      )
                                                      TEMP
                                                   )T
                                                   WHERE T.row_id = 2
                                                 )
                                                ORDER BY NLSSORT(name,'NLS_SORT= SCHINESE_PINYIN_M') asc,
                                                        job_number asc", yearAndMonthStr);
            System.Data.DataTable dtSameNamesButJN = OracleDaoHelper.getDTBySql(sqlStr);
            if (dtSameNamesButJN.Rows.Count == 0) {
                return;
            }
            //先确定文件路径
            string destFilePath = string.Format(@"{0}\考勤汇总\{1}_同名汇总_{2}.xls",System.Windows.Forms.Application.StartupPath,yearAndMonthStr,TimeHelper.getCurrentTimeStr());
            MyExcel myExcel = new MyExcel(destFilePath);
            myExcel.create();
            myExcel.openWithoutAlerts();
            //获取第一个Excel
            Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
            //确定行数目。
            int rowMaxCount = dtSameNamesButJN.Rows.Count * 2 + 6;
            //该月该考勤机共有多少个日期。
            int colMaxCount = AttendanceR.get_AR_Days_Num(YearAndMonthStr);
            #region  写记录
            //每行格式设置，注意标题占一行。
            Range range = firstSheet.get_Range(firstSheet.Cells[1, 1], firstSheet.Cells[rowMaxCount + 1, colMaxCount + 1]);
            //设置单元格为文本。
            range.NumberFormatLocal = "@";
            //水平对齐方式
            range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            //第一行写考勤分析结果。
            firstSheet.Cells[1, 1] = YearAndMonthStr + "异机同名用户考勤汇总" ;
            //获取该日期详细的考勤记录。
            #endregion
            #region
            V_AttendanceRecord.AR_Properties aR_Properties = V_AttendanceRecord.getARProperties(YearAndMonthStr);
            //第三行：考勤时间
            firstSheet.Cells[3, 1] = "考勤时间";
            firstSheet.Cells[3, 3] = String.Format(@"{0} ~ {1}",
                                                    aR_Properties.Start_Date,
                                                    aR_Properties.End_Date);
            firstSheet.Cells[3, 10] = "制表时间";
            firstSheet.Cells[3, 12] = aR_Properties.Tabulation_date;
            #endregion
            //获取该月该考勤机记录的考勤日期。
            List<int> dayList = V_AR_DETAIL.getARDaysOfSpecificMonth(YearAndMonthStr);
            lblPrompt.Text = string.Format(@"{0} 同名用户考勤汇总中：",YearAndMonthStr);
            lblPrompt.Visible = true;
            //写 此月与考勤相关的日。
            for (int i = 0; i <= dayList.Count - 1; i++)
            {
                //写该月的具体有哪些日：1，2，3.与考勤相关。
                firstSheet.Cells[4, i + 1] = dayList[i].ToString();
            }
            //实际出勤天数.
            firstSheet.Cells[4, dayList.Count + 1] = "实际出勤天数";
            //事假  
            firstSheet.Cells[4, dayList.Count + 2] = "事假";
            //未打卡
            firstSheet.Cells[4, dayList.Count + 3] = "未打卡";
            //平日加班
            firstSheet.Cells[4, dayList.Count + 4] = "平日延点";
            //加班日工作时长
            firstSheet.Cells[4, dayList.Count + 5] = "加班日工作时长";
            //加班合计
            firstSheet.Cells[4, dayList.Count + 6] = "加班合计(小时)";
            //迟到
            firstSheet.Cells[4, dayList.Count + 7] = "迟到";
            //早退
            firstSheet.Cells[4, dayList.Count + 8] = "早退";
            //餐补
            firstSheet.Cells[4, dayList.Count + 9] = "餐补";
            #region 填写刷卡细节
            //1.先填写 工号,姓名,部门.
            //1.工号:    3.工号    9.姓名:   11.姓名  19.部门:   21.部门.   
            for (int i=0;i<= dtSameNamesButJN.Rows.Count-1;i++) {
                //从第5行开始书写.
                firstSheet.Cells[5 + i * 2, 1] = "工 号:";
                firstSheet.Cells[5 + i * 2, 3] = dtSameNamesButJN.Rows[i]["job_number"].ToString();
                firstSheet.Cells[5 + i * 2, 9] = "姓 名:";
                firstSheet.Cells[5 + i * 2, 11] = dtSameNamesButJN.Rows[i]["name"].ToString();
                firstSheet.Cells[5 + i * 2, 19] = "部 门:";
                firstSheet.Cells[5 + i * 2, 21] = dtSameNamesButJN.Rows[i]["dept"].ToString();
            }
            //依据工号,书写考勤记录.
            for (int i = 0; i <= dtSameNamesButJN.Rows.Count - 1; i++)
            {
                string jn = dtSameNamesButJN.Rows[i]["job_number"].ToString();
                //首先依据工号,获取该月考勤细节.
                System.Data.DataTable dtARDetailByJN = getARDetailByJN(jn,yearAndMonthStr);
                for (int j=0;j<= dtARDetailByJN.Rows.Count-1;j++) {
                    //从第5行开始书写.
                    string fpt_first_time_str = dtARDetailByJN.Rows[j]["fpt_first_time"].ToString();
                    string fpt_last_time_str = dtARDetailByJN.Rows[j]["fpt_last_time"].ToString();
                    string come_late_num_str = dtARDetailByJN.Rows[j]["Come_Late_Num"].ToString();
                    string leave_early_num_str = dtARDetailByJN.Rows[j]["Leave_Early_Num"].ToString();
                    string tmpStr = fpt_first_time_str + "\r\n" + fpt_last_time_str;
                    firstSheet.Cells[6 + i * 2, 1 + j] = tmpStr;
                    if ("1".Equals(come_late_num_str)) {
                        ((Range)firstSheet.Cells[6 + i * 2, 1+j]).Characters[ 1, 5].Font.Color = -16776961;
                    }
                    if ("1".Equals(leave_early_num_str)) {
                        ((Range)firstSheet.Cells[6 + i * 2, 1 + j]).Characters[tmpStr.Length-4, 5].Font.Color = -16776961;
                    }
                }
            }
            #endregion
            int summar_start_col_Index = dayList.Count + 1;
            int summary_end_rowIndex = 4 + dtSameNamesButJN.Rows.Count * 2;
            #region 该写汇总.
            for (int i=0;i<=dtSameNamesButJN.Rows.Count-1;i++) {
                string jN = dtSameNamesButJN.Rows[i]["job_number"].ToString();
                V_Summary_OF_AR v_summary_of_ar = new V_Summary_OF_AR(jN, YearAndMonthStr);
                System.Data.DataTable dtARSummary = v_summary_of_ar.getSummaryOFAR();

                //实际出勤天数.
                firstSheet.Cells[6 + i*2, summar_start_col_Index] = dtARSummary.Rows[0]["AR_DAYS"].ToString();

                //事假 
                string vacatioin_total_time = dtARSummary.Rows[0]["VACATION_TOTAL_TIME"].ToString();
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 1] = "0".Equals(vacatioin_total_time) ? "" : vacatioin_total_time;

                string not_Finger_Print_num = dtARSummary.Rows[0]["NOT_FINGERPRINT_TIMES"].ToString();
                //未打卡
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 2] = "0".Equals(not_Finger_Print_num) ? "" : not_Finger_Print_num;
                //平日延时
                string DELAY_TIMES_OF_Ordinary_Days_str = dtARSummary.Rows[0]["DELAY_TIMES_OF_Ordinary_Days"].ToString();
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 3] = "0.0".Equals(DELAY_TIMES_OF_Ordinary_Days_str) ? "" : DELAY_TIMES_OF_Ordinary_Days_str;
                //写加班日工作时长
                string duration_ov_overtime_days_str = dtARSummary.Rows[0]["Duration_Of_Overtime_Days"].ToString();
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 4] = "0.0".Equals(duration_ov_overtime_days_str) ? "" : duration_ov_overtime_days_str;
                //写总的加班费用。
                string delayTotalTimes_Str = dtARSummary.Rows[0]["DELAY_TOTAL_TIME"].ToString();
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 5] = "0.0".Equals(delayTotalTimes_Str) ? "" : delayTotalTimes_Str;

                string come_late_Num = dtARSummary.Rows[0]["COME_LATE_NUM"].ToString();
                //迟到
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 6] = "0".Equals(come_late_Num) ? "" : come_late_Num;
                string leave_early_num = dtARSummary.Rows[0]["LEAVE_EARLY_NUM"].ToString();
                //早退
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 7] = "0".Equals(leave_early_num) ? "" : leave_early_num;
                //餐补
                firstSheet.Cells[6 + i * 2, summar_start_col_Index + 8] = dtARSummary.Rows[0]["DINNER_SUBSIDY_NUM"].ToString();
            }
            #endregion
            #region  标注加班日
            //休息日，背景色变为浅绿色。
            for (int colIndex = 1; colIndex <= dayList.Count; colIndex++)
            {
                bool iftheDaysOfOverTime = false;
                string dayStr = ((Range)(firstSheet.Cells[4, colIndex])).Text.ToString();
                string AR_Day = YearAndMonthStr + dayStr.PadLeft(2, '0');
                iftheDaysOfOverTime = Have_A_Rest_Helper.ifDayOfRest(AR_Day);
                if (iftheDaysOfOverTime)
                {
                    Range rangetheDaysOfOverTime = firstSheet.get_Range(firstSheet.Cells[4, colIndex], firstSheet.Cells[summary_end_rowIndex, colIndex]);
                    rangetheDaysOfOverTime.Interior.Pattern = XlPattern.xlPatternSolid;
                    rangetheDaysOfOverTime.Interior.PatternColorIndex = XlPattern.xlPatternAutomatic;
                    rangetheDaysOfOverTime.Interior.ThemeColor = XlThemeColor.xlThemeColorAccent3;
                    rangetheDaysOfOverTime.Interior.TintAndShade = 0.599993896298105;
                    rangetheDaysOfOverTime.Interior.PatternTintAndShade = 0;
                }
            }
            #endregion
            //合并第一行
            Range rangeTitle = firstSheet.get_Range(firstSheet.Cells[1, 1], firstSheet.Cells[1, summar_start_col_Index + 8]);
            rangeTitle.Merge();
            rangeTitle.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            rangeTitle.VerticalAlignment = XlVAlign.xlVAlignCenter;
            Range rangToReplace = firstSheet.get_Range(firstSheet.Cells[6, summar_start_col_Index], firstSheet.Cells[summary_end_rowIndex, summar_start_col_Index + 8]);
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstSheet);
            uEHelper.replace_Str_in_the_range(rangToReplace, "0", "");
            //对汇总区域进行列宽自适应。
            myExcel.save();
            myExcel.close();
            lblPrompt.Text = string.Format(@"{0} 同名用户考勤汇总结束：", YearAndMonthStr);
            lblPrompt.Visible = true;
        }
        */
        /*
        /// <summary>
        /// 依据工号,获取考勤细节.
        /// </summary>
        /// <param name="jn"></param>
        /// <returns></returns>
        private System.Data.DataTable getARDetailByJN(string jn,string year_and_month_str)
        {
            string sqlStr = string.Format(@"Select 
                                              SUBSTR(to_char(fpt_first_time,'yyyy-MM-dd HH24:MI:SS'),12,5) fpt_first_Time,
                                              SUBSTR(to_char(fpt_last_time,'yyyy-MM-dd HH24:MI:SS'),12,5) fpt_last_time,
                                              CAST(come_late_num as varchar2(2)) AS Come_Late_Num,
                                              CAST(leave_early_num as varchar2(2)) AS Leave_Early_Num
                                        from Attendance_Record
                                        where job_number = '{0}'
                                        and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')
                                        order by fingerprint_date asc",
                                        jn,
                                        year_and_month_str);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }
        */
        private void lblResult_Click(object sender, EventArgs e)
        {
            int hwnd = 0;
            try
            {
                hwnd = ExcelHelper.openBook(destFilePath);
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
            timerShowProgress.Enabled = false;
            NotifyIcon nI = ((FrmMainOfAttendanceRecord)this.MdiParent).nfiSystem;
            if (nI.Visible) {
                nI.ShowBalloonTip(20000, "提示：",lblPrompt.Text.Trim() + ":   "+ Math.Round(((decimal)(pb.Value/pb.Maximum) * 100),2).ToString(), ToolTipIcon.Info);
            }
            timerShowProgress.Enabled = true;
        }

        private void timerCompleted_Tick(object sender, EventArgs e)
        {
            timerCompleted.Stop();
            try {
                NotifyIcon nI = ((FrmMainOfAttendanceRecord)this.MdiParent).nfiSystem;
                if (_status)
                    nI.Icon = System.Drawing.Icon.ExtractAssociatedIcon(path + "blank.ico");
                else
                    nI.Icon = Properties.Resources.apps;
                _status = !_status;
            }
            catch (Exception ex) {
                MessageBox.Show(ex.ToString());
            }
            timerCompleted.Start();
        }
        private void FrmAnalyzeAR_Load(object sender, EventArgs e)
        {
            if ("AttendanceRecord".Equals(System.Windows.Forms.Application.ProductName))
            {
                path =  System.Windows.Forms.Application.StartupPath+ "\\";
            }
            else {
                path = Environment.CurrentDirectory + "\\" + System.Windows.Forms.Application.ProductVersion + "\\";
            }
            YearAndMonthStr = mCalendar.SelectionRange.Start.ToString("yyyy-MM");
        }
        private void cb_Attendance_Machine_SelectedIndexChanged(object sender, EventArgs e)
        {
        }
        /*
        /// <summary>
        /// 
        /// </summary>
        /// <param name="yearAndMonthStr"></param>
        /// <returns></returns>
        private System.Data.DataTable getSameNamesButDifferentMachineNo(string yearAndMonthStr) {
            string sqlStr = string.Format(@"SELECT distinct dept,job_number,name
                                                FROM Attendance_Record AR
                                                WHERE TRUNC(Fingerprint_date,'MM') = TO_DATE('{0}','yyyy-MM')
                                                AND AR.name = ANY
                                                (
                                                  SELECT name 
                                                   FROM  
                                                   (
                                                      SELECT row_number() over(partition by name order by job_number asc) row_id,
                                                             job_number,
                                                             name               
                                                      FROM 
                                                      (
                                                           SELECT DISTINCT name,
                                                                      job_number
                                                           from Attendance_Record AR
                                                           where trunc(fingerprint_date,'MM') = to_date('{0}','yyyy-MM')
                                                      )
                                                      TEMP
                                                   )T
                                                   WHERE T.row_id = 2
                                                 )
                                                ORDER BY NLSSORT(name,'NLS_SORT= SCHINESE_PINYIN_M') asc,
                                                        job_number asc", yearAndMonthStr);
            System.Data.DataTable dtSameNamesButJN = OracleDaoHelper.getDTBySql(sqlStr);
            return dtSameNamesButJN;
        }
        */
    }
}
