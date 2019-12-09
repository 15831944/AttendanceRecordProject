using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Tools;

namespace AttendanceRecord.Entities
{
    public class V_Work_Schedule
    {
        private string _DEPT;
        private string _ON_DUTY_TIME;
        private string _RECORD_TIME;
        private string _WORK_AND_REST_DATE;
        private string _WORK_OR_REST;
        private string _OFF_DUTY_TIME;
        private string _FP_NUMBER;
        private string _rest_number;
        public static string _YearAndMonthStr = String.Empty;
        public static Queue<int> hwndOfXls_Queue = new Queue<int>();
        public string DEPT
        {
            get
            {
                return _DEPT;
            }

            set
            {
                _DEPT = value;
            }
        }

        public string ON_DUTY_TIME
        {
            get
            {
                return _ON_DUTY_TIME;
            }

            set
            {
                _ON_DUTY_TIME = value;
            }
        }

        public string RECORD_TIME
        {
            get
            {
                return _RECORD_TIME;
            }

            set
            {
                _RECORD_TIME = value;
            }
        }

        public string WORK_AND_REST_DATE
        {
            get
            {
                return _WORK_AND_REST_DATE;
            }

            set
            {
                _WORK_AND_REST_DATE = value;
            }
        }

        public string WORK_OR_REST
        {
            get
            {
                return _WORK_OR_REST;
            }

            set
            {
                _WORK_OR_REST = value;
            }
        }

        public string OFF_DUTY_TIME
        {
            get
            {
                return _OFF_DUTY_TIME;
            }

            set
            {
                _OFF_DUTY_TIME = value;
            }
        }

        public string FP_NUMBER
        {
            get
            {
                return _FP_NUMBER;
            }

            set
            {
                _FP_NUMBER = value;
            }
        }

        public string Rest_number
        {
            get
            {
                return _rest_number;
            }

            set
            {
                _rest_number = value;
            }
        }
        #region 给予日期
        public int GenWorkSchedule()
        {
            //已经存在某个制定月的工作计划，先删除。
            if (ifExistsWS())
            {
                delWS();
            }
            string sqlStr = String.Format(@"INSERT INTO Work_Schedule(
                                                SEQ,
                                                DEPT,
                                                ON_DUTY_TIME,
                                                RECORD_TIME,
                                                WORK_AND_REST_DATE,
                                                WORK_OR_REST,
                                                OFF_DUTY_TIME,
                                                FP_NUMBER,
                                                REST_NUMBER
                                            )
                                            SELECT 
                                                SEQ_Work_Schedule.nextval,
                                                DEPT,
                                                ON_DUTY_TIME,
                                                RECORD_TIME,
                                                WORK_AND_REST_DATE,
                                                WORK_OR_REST,
                                                OFF_DUTY_TIME,
                                                FP_NUMBER,
                                                REST_NUMBER    
                                            FROM Work_Summary
                                            WHERE TRUNC(WORK_AND_REST_DATE,'MM') = TO_DATE('{0}','YYYY-MM')",
                                            _YearAndMonthStr);
            return Tools.OracleDaoHelper.executeSQL(sqlStr);
        }
        #endregion
        public System.Data.DataTable getWorkSchedule()
        {
            string sqlStr = String.Format(@"
                                            SELECT 
                                               dept, 
                                                work_and_rest_date, 
                                                on_duty_time, 
                                                off_duty_time, 
                                                work_or_rest, 
                                                fp_number, 
                                                rest_number
                                            FROM Work_Schedule
                                            WHERE TRUNC(WORK_AND_REST_DATE,'MM') = TO_DATE('{0}','YYYY-MM')", V_Work_Schedule._YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        }
        #region 获取工作计划列表。
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<V_Work_Schedule> getWorkScheduleList()
        {
            string sqlStr = String.Format(@"
                                            SELECT 
                                               dept, 
                                                work_and_rest_date, 
                                                on_duty_time, 
                                                off_duty_time, 
                                                work_or_rest, 
                                                fp_number, 
                                                rest_number
                                            FROM Work_Schedule
                                            WHERE TRUNC(WORK_AND_REST_DATE,'MM') = TO_DATE('{0}','YYYY-MM')", V_Work_Schedule._YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return ConvertHelper<V_Work_Schedule>.ConvertToList(dt);
        }
        #endregion
        #region 生成Excel
        public Tools.MSG genExcel(out string xlsFilePath)
        {
            killHwndOfXls();
            MSG msg = new MSG();
            string excelName = _YearAndMonthStr + "_工作安排表.xls";
            string defaultDir = Environment.CurrentDirectory + string.Format(@"\工作安排表", excelName);
            string xlsFileFullPath = FileNameDialog.getSaveFileNameWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir, excelName);
            if (!xlsFileFullPath.Contains(@"\"))
            {
                msg.Flag = false;
                msg.Msg = "取消了选择！";
                xlsFilePath = xlsFileFullPath;
                return msg;
            }
            ApplicationClass app = new ApplicationClass();
            hwndOfXls_Queue.Enqueue(app.Hwnd);
            app.Visible = false;
            Workbook wBook = null;
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            
            try
            {
                wBook = app.Workbooks.Add(true);
                Worksheet wS = wBook.Worksheets.Item[1] as Worksheet;
                //每行格式设置，注意标题占一行。
                rowsMaxCount = this.getDeptNum() + 4;
                Queue<int> daysQueue = this.getAllDaysOfThisMonth();
                colsMaxCount = daysQueue.Count;
                //每行格式设置，注意标题占一行。
                Range range = wS.get_Range(wS.Cells[1, 1], wS.Cells[rowsMaxCount + 1, colsMaxCount + 1]);
                //设置单元格为文本。
                range.NumberFormatLocal = "@";
                //水平对齐方式
                range.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                //第一行写考勤分析结果。
                wS.Cells[1, 1] = _YearAndMonthStr + " 工作安排分析结果";
                wS.Cells[1, 3] = "休假";
                ((Range)(wS.Cells[1, 4])).Interior.Color = System.Drawing.Color.Green.ToArgb();                //获取该日期详细的考勤记录。
                wS.Cells[1, 5] = "周日";
                ((Range)wS.Cells[1, 6]).Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                V_AR_DETAIL v_AR_Detail = V_AR_DETAIL.get_V_AR_Detail(_YearAndMonthStr);
           
                //第三行：考勤时间
                wS.Cells[3, 1] = "考勤时间";
                wS.Cells[3, 3] = String.Format(@"{0} ~ {1}",
                                                        v_AR_Detail.Start_date,
                                                        v_AR_Detail.End_date);
                wS.Cells[3, 10] = "制表时间";
                wS.Cells[3, 12] = v_AR_Detail.Tabulation_time;
                
                //书写部门
                List<string> deptList = this.getDeptList();
                for (int i = 0; i <= deptList.Count - 1; i++) {
                    wS.Cells[5+i,1] = deptList[i].ToString();
                }
                string v_w_s_Day = string.Empty;
                //开始书写此月的日子
                int colIndex = 1;
                while (daysQueue.Count>0) {
                    colIndex++;
                    int day = daysQueue.Dequeue();
                    wS.Cells[4, colIndex] = day;
                    v_w_s_Day = _YearAndMonthStr + "-" + day.ToString().PadLeft(2, '0');
                    for (int i = 0; i <= deptList.Count - 1; i++)
                    {
                        //填写该日内所有部门的休假情况.
                        V_W_S v_W_S = new V_W_S();
                        v_W_S.Dept = deptList[i].ToString();
                        v_W_S.Work_and_rest_date = v_w_s_Day;
                        v_W_S = v_W_S.get_V_W_S_By_Date_And_Dept();
                        wS.Cells[5+i, colIndex] = v_W_S.Work_rate;
                        if ("休息".Equals(v_W_S.Work_or_rest, StringComparison.OrdinalIgnoreCase)) {
                            ((Range)wS.Cells[5 + i, colIndex]).Interior.Color = System.Drawing.Color.Green.ToArgb();
                            continue;
                        }
                        //注意周日为每周的第一天。
                        if ("1" == v_W_S.Day_of_week)
                            //周日用暗灰色
                            ((Range)wS.Cells[5 + i, colIndex]).Interior.Color = System.Drawing.Color.Yellow.ToArgb();
                    }
                   
                }
                //自动调整列宽
                //range.EntireColumn.AutoFit();
                //设置禁止弹出保存和覆盖的询问提示框
                app.DisplayAlerts = false;
                app.AlertBeforeOverwriting = false;
                //保存excel文档并关闭
                wBook.SaveAs(xlsFileFullPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                wBook.Close(true, xlsFileFullPath, Type.Missing);
                //退出Excel程序
                app.Quit();
                //释放资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wS);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                //调用GC的垃圾收集方法
                GC.Collect();
                GC.WaitForPendingFinalizers();
                msg.Msg = "工作安排: " + xlsFileFullPath;
                msg.Flag = true;
                xlsFilePath = xlsFileFullPath;
                return msg;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "提示消息:", MessageBoxButtons.OK, MessageBoxIcon.Information);
                msg.Msg = ex.ToString();
                msg.Flag = false;
                xlsFilePath = xlsFileFullPath;
                return msg;
            }
        }
        #endregion
        /// <summary>
        /// 是否存在制定日期的工作计划。
        /// </summary>
        /// <returns></returns>
        public bool ifExistsWS()
        {
            bool result = false;
            string sqlStr = String.Format(@"
                                            select 1 
                                            from Work_Schedule
                                            where trunc(work_and_rest_date,'MM') = To_DATE('{0}','yyyy-MM')
                                                ", V_Work_Schedule._YearAndMonthStr);
            result = OracleDaoHelper.getDTBySql(sqlStr).Rows.Count > 0 ? true : false;
            return result;
        }
        /// <summary>
        /// 删除制定日期的工作计划。
        /// </summary>
        /// <returns></returns>
        public int delWS()
        {
            string sqlStr = String.Format(@"
                                            delete
                                            from Work_Schedule
                                            where trunc(work_and_rest_date,'MM') = To_DATE('{0}','yyyy-MM')",
                                            V_Work_Schedule._YearAndMonthStr);
            return OracleDaoHelper.executeSQL(sqlStr);
        }
        #region 获取有多少个部门
        public int getDeptNum()
        {
            string sqlStr = String.Format(@"select DISTINCT(Dept) Dept 
                                            from v_Work_Schedule
                                           where trunc(work_and_rest_date,'MM') = To_DATE('{0}','yyyy-MM')
                                            order by dept asc", V_Work_Schedule._YearAndMonthStr);
            System.Data.DataTable dt = (System.Data.DataTable)(OracleDaoHelper.getDTBySql(sqlStr));
            return dt.Rows.Count;
        }
        #endregion;
        #region 获取所有部门
        public List<string> getDeptList() {
            string sqlStr = String.Format(@"select DISTINCT(Dept) Dept 
                                            from v_Work_Schedule
                                           where trunc(work_and_rest_date,'MM') = To_DATE('{0}','yyyy-MM')
                                           order by dept asc", V_Work_Schedule._YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            List<String> deptList = new List<string>();
            for (int i=0;i<=dt.Rows.Count-1;i++) {
                deptList.Add(dt.Rows[i]["Dept"].ToString());
            }
            return deptList;
        }
        #endregion
        #region 获取有多少日，此月
        public Queue<int> getAllDaysOfThisMonth()
        {
            string sqlStr = String.Format(@"SELECT DISTINCT(SUBSTR(TO_CHAR(work_and_Rest_Date,'YYYY-MM-DD'),9,2)) AS Day
                                            FROM V_Work_Schedule
                                            where trunc(work_and_rest_date, 'MM') = To_DATE('{0}', 'yyyy-MM')
                                            ORDER BY Day ASC",
                                            V_Work_Schedule._YearAndMonthStr);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            Queue<int> AllDaysOfThisMonthQ = new Queue<int>();
            for (int i=0;i<=dt.Rows.Count-1;i++) {
                AllDaysOfThisMonthQ.Enqueue(int.Parse(dt.Rows[i]["Day"].ToString()));
            }
            return AllDaysOfThisMonthQ;
        }
        #endregion;
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

    }
}
