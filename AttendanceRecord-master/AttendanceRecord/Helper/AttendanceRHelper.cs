using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Tools;
using AttendanceRecord.Entities;
using Excel;
using Oracle.DataAccess.Client;
using System.Data;
using System.ComponentModel;
namespace AttendanceRecord.Helper
{
    public  class AttendanceRHelper
    {
        public static string tempStr = String.Empty;
        //导入数据的行数.
        public static int affectedCount = 0;
        /// <summary>
        /// 将考勤记录导入数据库.
        /// </summary>
        /// <param name="xlsFilePath"></param>
        /// <param name="randomStr"></param>
        /// <param name="pb"></param>
        /// <returns></returns>
        public static MSG  ImportAttendanceRecordToDB(string xlsFilePath, string randomStr, BackgroundWorker bgWork)
        {
            string excelName = Usual_Excel_Helper.getExcelName(xlsFilePath);
            bgWork.ReportProgress(0, string.Format(@"lblPrompt.Text = {0}，准备读取：", excelName));
            int pbLength = 0;
            bgWork.ReportProgress(pbLength, "pb.Maximum");
            int pbValue = 0;
            bgWork.ReportProgress(pbValue, "pb.Value");
            MSG msg = new MSG();

            //用于确定本月最后一天.
            Stack<int> sDate = new Stack<int>();
            //Queue<AttendanceR> qAttendanceR = new Queue<AttendanceR>();
            Queue<AttendanceRecordDetail> qARDetail = new Queue<AttendanceRecordDetail>();
            AttendanceRecordDetail._random_str = randomStr;
            //按指纹日期
            string fingerPrintDate = String.Empty;
           
            //行最大值.
            int rowsMaxCount = 0;
            int colsMaxCount = 0;
            Usual_Excel_Helper uEHelper = null;
           
            MyExcel myExcel = new MyExcel(xlsFilePath);
            //打开该文档。
            myExcel.openWithoutAlerts();
            //只获取第一个表格。
            Worksheet ws = myExcel.getFirstWorkSheetAfterOpen();
            bgWork.ReportProgress(0, string.Format(@"lblPrompt.Text = {0}，正在读取：", excelName));
            AttendanceRecordDetail._file_path = xlsFilePath;
            //行;列最大值 赋值.
            rowsMaxCount = ws.UsedRange.Rows.Count;
            colsMaxCount = ws.UsedRange.Columns.Count;


            AttendanceRecordDetail._sheet_name = ws.Name;
            //判断首行是否为 考勤记录表;以此判断此表是否为考勤记录表.
            string A1Str = ((Range)ws.Cells[1, 1]).Text.ToString().Trim().Replace("\n", "").Replace("\r", "").Replace(" ", "");
            if (String.IsNullOrEmpty(A1Str))
            {
                msg.Msg = "工作表的A1单元格不能为空！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //如果A1Str的内容不包含"考勤记录表"5个字。       
            if (!A1Str.Contains("考勤记录表"))
            {
                msg.Msg = "A1内容未包含'考勤记录表'";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #region 判断名称中是否区分了考勤记录。
            string Seq_Attendance_Record = string.Empty;
            int indexOfFullStop = xlsFilePath.LastIndexOf(".");
            Seq_Attendance_Record = xlsFilePath.Substring(indexOfFullStop - 1, 1);
            if (!CheckPattern.CheckNumber(Seq_Attendance_Record))
            {
                msg.Msg = "考勤记录表名称请以数字结尾！";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            #endregion
            
            AttendanceRecordDetail._prefix_Job_Number = excelName.Substring(excelName.Length - 1, 1).ToCharArray()[0];
            string C3Str = ((Range)ws.Cells[3, 3]).Text.ToString().Trim();
            //  \0: 表空字符.
            if (String.IsNullOrEmpty(C3Str))
            {
                msg.Msg = "异常: 考勤时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //
            string[] ArrayC3 = C3Str.Split('~');
            if (ArrayC3.Length == 0)
            {
                msg.Msg = "异常: 考勤时间格式变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            AttendanceRecordDetail._start_date = ArrayC3[0].ToString().Trim().Replace('/', '-');
            AttendanceRecordDetail._end_date = ArrayC3[1].ToString().Trim().Replace('/', '-');
            //制表时间:  L3 3行12列.
            string L3Str = ((Range)ws.Cells[3, 12]).Text.ToString().Trim().Replace('/', '-');
            if (String.IsNullOrEmpty(L3Str))
            {
                msg.Msg = "异常: 制表时间为空!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            //制表时间.
            AttendanceRecordDetail._tabulation_time = L3Str;
            //检查第4行是否为;考勤时间:
            string A4Str = ((Range)ws.Cells[4, 1]).Text.ToString().Trim();
            if (!"1".Equals(A4Str, StringComparison.CurrentCultureIgnoreCase))
            {
                msg.Msg = "异常: 第四行已变更!";
                msg.Flag = false;
                myExcel.close();
                return msg;
            }
            uEHelper = new Usual_Excel_Helper(ws);
            //此刻不能删除，只是获取行号。
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            //判断是否有空行。
            for (int i = 5; i <= rowsMaxCount; i++)
            {
                if (uEHelper.isBlankRow(i))
                {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp))
                    {
                        //本行为空，上一行为工号行，则也统计。
                        continue;
                    }
                    //本行，为空，上一行非工号行。则删除本行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                };
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count > 0)
            {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            };
            rowsMaxCount = ws.UsedRange.Rows.Count;
            //进度条长度增加。
            pbLength += colsMaxCount;
            pbLength += (colsMaxCount * (rowsMaxCount - 5 + 1));
            bgWork.ReportProgress(pbLength, "pb.Maximum");
            //入队列值0
            sDate.Push(0);
            //显示进度条。
            //考勤表中第4行，某月的最大考勤天数。
            //lblPrompt.Text = excelName + "，正在读取：";
            
            int actualMaxDay = 0;
            //开始循环
            for (int i = 1; i <= colsMaxCount; i++)
            {
                A4Str = ((Range)ws.Cells[4, i]).Text.ToString();
                //碰到第4行某列为空，退出循环。
                if (String.IsNullOrEmpty(A4Str))
                {
                    break;
                }
                int aDate = 0;
                //对A4Str进行分析.
                if (!Int32.TryParse(A4Str, out aDate))
                {
                    msg.Msg = String.Format(@"异常: 考勤日期行第{0}列出现非数字内容!", aDate);
                    msg.Flag = false;
                    myExcel.close();
                    return msg;
                }
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek())
                {
                    //跳出循环.
                    break;
                }
                actualMaxDay++;
                sDate.Push(aDate);
                //pb.Value++;
                bgWork.ReportProgress(pbValue++, "pb.Value");
            }
            //取其中的最小值。
            colsMaxCount = Math.Min(sDate.Count - 1, actualMaxDay);
            //考勤日期
            fingerPrintDate = AttendanceRecordDetail._start_date.Substring(0, 7).Replace('/', '-');
            string tempStr = string.Empty;
            //开始循环
            for (int colIndex = 1; colIndex <= colsMaxCount; colIndex++)
            {
                //从第5行开始.
                //奇数;偶数行共用一个对象.
                AttendanceRecordDetail ARDetail = null;
                //设定用于填充的对象
                AttendanceRecordDetail._prefix_Job_Number = Seq_Attendance_Record.ToCharArray()[0];
                for (int rowIndex = 5; rowIndex <= rowsMaxCount; rowIndex++)
                {
                    //如果行数为奇数则为工号行.
                    if (rowIndex % 2 == 1)
                    {
                        //工号行.
                        //取工号
                        ARDetail = new AttendanceRecordDetail();
                        ARDetail.Job_number = ((Range)ws.Cells[rowIndex, 3]).Text.ToString().Trim();
                        //自行拼凑AR.
                        ARDetail.combine_Job_Number();
                        //取姓名:  K5 
                        ARDetail.Name = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("K")]).Text.ToString().Trim();
                        //取部门: U5
                        ARDetail.Dept = ((Range)ws.Cells[rowIndex, Usual_Excel_Helper.getColIndexByStr("U")]).Text.ToString().Trim();
                        //部门为空，则填充为NULL;
                        ARDetail.Dept = !String.IsNullOrEmpty(ARDetail.Dept) ? ARDetail.Dept : "NULL";
                        //取日期.填充0;
                        ARDetail.Fingerprint_date = fingerPrintDate + "-" + colIndex.ToString().PadLeft(2, '0');
                    }
                    else
                    {
                        //偶数行取考勤结果.
                        //上班时间. 如B10;
                        tempStr = ((Range)ws.Cells[rowIndex, colIndex]).Text.ToString().Trim();
                        string tempFirstTime = String.Empty;
                        string tempLastTime = String.Empty;
                        List<string> strTimeList = null;
                        msg = getFPTimeReturnMSG(tempStr, out strTimeList);
                        if (!msg.Flag)
                        {
                            msg.Msg = string.Format(@"导入失败,提交数据尚未开始：第{0}行{1}列,{1}！", rowIndex, colIndex, msg.Msg);
                            myExcel.close();
                            return msg;
                        };
                        //无打卡记录,不提交
                        if (strTimeList.Count==0)
                        {
                            qARDetail.Enqueue(ARDetail);
                        }
                        //有打卡记录
                        for (int i = 0; i < strTimeList.Count; i++)
                        {
                            AttendanceRecordDetail ARDetailTemp = (AttendanceRecordDetail)CloneObject.Clone(ARDetail);
                            ARDetailTemp.Finger_print_time = ARDetailTemp.Fingerprint_date + " " + strTimeList[i].ToString();
                            qARDetail.Enqueue(ARDetailTemp);
                        }
                    }
                    //pb.Value++;
                    bgWork.ReportProgress(pbValue++, "pb.Value");
                }
            }
            //释放对象
            myExcel.close();
            System.Threading.Thread.Sleep(2000);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //lblResult.Text = "";
            bgWork.ReportProgress(0, "lblResult.Text = ''");
            //lblPrompt.Text = "提交数据: ";
            bgWork.ReportProgress(0, string.Format(@"lblPrompt.Text = {0}, 提交数据:",excelName));
            //
            bgWork.ReportProgress(qARDetail.Count, "pb.Maximum");
            //*******/
            pbValue = 0;
            bgWork.ReportProgress(pbValue, "pb.Value");
            #region
            //OracleDaoHelper.noLogging("Attendance_Record");
            OracleDaoHelper.noLogging("Attendance_Record_Detail");
            OracleConnection conn = OracleConnHelper.getConn();
            OracleTransaction tran = conn.BeginTransaction();
            //保存对象
            while (qARDetail.Count > 0)
            {
                try
                {
                    AttendanceRecordDetail aRDetail = qARDetail.Dequeue();
                 
                    affectedCount += aRDetail.saveBySpecificConn(conn);
                    //pb.Value++;
                    bgWork.ReportProgress(pbValue++, "pb.Value");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "提示:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    msg.Msg = DirectoryHelper.getFileName(xlsFilePath) + "：导入失败; " + ex.ToString();
                    msg.Flag = false;
                    tran.Rollback();
                    conn.Close();
                    conn.Dispose();
                    return msg;
                    throw;
                }
            }
            tran.Commit();
            conn.Close();
            conn.Dispose();
            #endregion
            //OracleDaoHelper.logging("Attendance_Record");
            OracleDaoHelper.logging("Attendance_Record_Detail");
            msg.Flag = true;
            msg.Msg = String.Format(@"导入完成;总计{0}条.", affectedCount.ToString());
            return msg;
        }

        private static void NewMethod(BackgroundWorker bgWork, int pbLength)
        {
            bgWork.ReportProgress(pbLength, "pb.Maximum");
        }

        /// <summary>
        /// 将考勤记录导入预备表中.
        /// </summary>
        /// <param name="xlsFilePath"></param>
        /// <param name="randomStr"></param>
        /// <param name="pb"></param>
        /// <returns></returns>
        /// 
        /// <summary>
        /// 目前只考虑白班.
        /// </summary>
        /// <param name="strTime"></param>
        /// <param name="FirstTime"></param>
        /// <param name="lastTime"></param>
        public static bool GetFPTime(string strTime, out string firstTime, out string lastTime)
        {
            int hour = 0;
            int minute = 0;
            //目的是再次到处汇总的表格，因为汇总的表格中时间纪录以 \r\n 换行相隔。
            strTime = strTime.Replace("\r\n", "");
            strTime = strTime.Replace(" ", "");
            strTime = strTime.Replace("  ", "");
            strTime = strTime.Replace("   ", "");
            //判断长度
            if (String.IsNullOrEmpty(strTime))
            {
                firstTime = "";
                lastTime = "";
                return true;
            }
            if (strTime.Substring(1, 1) == ":")
            {
                //如7:07 --> 07:07
                strTime = "0" + strTime;
            }
            //判断长度是否可以被5整除
            if (strTime.Length % 5 != 0)
            {
                firstTime = "";
                lastTime = "";
                return false;
            }
            List<string> timeStrList = new List<string>();
            for (int i = 0; i <= strTime.Length / 5 - 1; i++)
            {
                timeStrList.Add(strTime.Substring(i * 5, 5));
            }
            //排序好的字符串。
            List<DateTime> dtSortedList = new List<DateTime>();
            for (int i = 0; i <= timeStrList.Count - 1; i++)
            {
                //如果 时间字符串格式不符合规定，return false;
                bool flag = false;
                DateTime dt;
                flag = DateTime.TryParse(timeStrList[i], out dt);
                if (!flag)
                {
                    firstTime = "";
                    lastTime = "";
                    return false;
                }
                dtSortedList.Add(dt);
            }
            //排序，再次输出。
            dtSortedList.Sort();
            strTime = string.Empty;
            for (int i = 0; i <= dtSortedList.Count - 1; i++)
            {
                strTime += dtSortedList[i].ToString("HH:mm");

            }
            #region 如果时间字符长度为5; 表示只刷了一次卡.
            if (strTime.Length == 5)
            {
                //判断是否为:  6 --> 11点
                hour = 0;
                minute = 0;
                //判断是否>=0:01 And< 12:20  
                int.TryParse(strTime.Substring(0, 2), out hour);
                int.TryParse(strTime.Substring(3, 2), out minute);
                if (hour >= 6 && (hour < 11))
                {
                    firstTime = strTime;
                    lastTime = "";
                    return true;
                }
                else if (hour < 6 && hour >= 0)
                {
                    //凌晨间的刷卡，也 计为：早上卡。
                    firstTime = strTime;
                    lastTime = "";
                    return true;
                }
                else
                {
                    firstTime = "";
                    lastTime = strTime;
                    return true;
                }
            }
            #endregion
            #region 刷卡次数为2次，判断第一次和第二次的间隔时间是否小于10分钟。
            //若小于10分钟，则认为Fpt_last_time, 不填写。
            if (10 == strTime.Length)
            {
                //获取
                string str1 = strTime.Substring(0, 5);
                string str2 = strTime.Substring(5, 5);

                DateTime dt1 = DateTime.Parse(str1);
                DateTime dt2 = DateTime.Parse(str2);
                double differentValue = (dt2 - dt1).TotalMinutes;
                if (differentValue < 10)
                {
                    firstTime = str1;
                    lastTime = "";
                    return true;
                }
                firstTime = str1;
                lastTime = str2;
                return true;
            }
            #endregion
            #region 刷卡次数:  2次以上.
            firstTime = strTime.Substring(0, 5);
            lastTime = strTime.Substring(strTime.Length - 5, 5);
            #endregion
            return true;

        }

        #region 目的是获取打卡时间字符串。
        public static MSG getFPTimeReturnMSG(string strTime,out List<String> outTimeStrList)
            {
                MSG msg = new MSG("", false);
                List<String> timeStrList = new List<string>();
                //目的是再次到处汇总的表格，因为汇总的表格中时间纪录以 \r\n 换行相隔。
                strTime = strTime.Replace("\r\n", "");
                strTime = strTime.Replace(" ", "");
                strTime = strTime.Replace("  ", "");
                strTime = strTime.Replace("   ", "");
                //判断长度
                if (String.IsNullOrEmpty(strTime))
                {
                    outTimeStrList = timeStrList;
                    msg.Flag = true;
                    return msg;
                }
                if (strTime.Substring(1, 1) == ":")
                {
                    //转换 7:07 --> 07:07
                    strTime = "0" + strTime;
                }
                //判断长度是否可以被5整除
                if (strTime.Length % 5 != 0)
                {
                    outTimeStrList = timeStrList;
                    msg.Msg = "时间字符串长度不能被5整除！";
                    return msg;
                }
                for (int i = 0; i <= strTime.Length / 5 - 1; i++)
                {
                    timeStrList.Add(strTime.Substring(i * 5, 5));
                }
                //排序好的字符串。
                List<DateTime> dtSortedList = new List<DateTime>();
                for (int i = 0; i <= timeStrList.Count - 1; i++)
                {
                    //如果 时间字符串格式不符合规定，return false;
                    bool flag = false;
                    DateTime dt;
                    flag = DateTime.TryParse(timeStrList[i], out dt);
                    if (!flag)
                    {
                        outTimeStrList = timeStrList;
                        msg.Msg = "字符串以长度5转换为数组后，检查后存在非时间的字符串：" + timeStrList[i].ToString();
                        return msg;
                    }
                    dtSortedList.Add(dt);
                }
                //排序，再次输出。
                dtSortedList.Sort();
                string strTimeTemp = string.Empty;
                //清空 timeStrList
                timeStrList.Clear();
                for (int i = 0; i <= dtSortedList.Count - 1; i++)
                {
                    strTimeTemp = dtSortedList[i].ToString("HH:mm");
                    timeStrList.Add(strTimeTemp);
                }
                outTimeStrList = timeStrList;
                msg.Msg = "正常！";
                msg.Flag = true;
                return msg;
            }
            #endregion

            #region 检查时间内容是否正确
            public static MSG checkTimeStr(string timStr) {
                MSG msg = new MSG();
                //判断是否为5的倍数.
                if (timStr.Length % 5 != 0) {
                    msg.Flag = false;
                    msg.Msg = "时间长度不为5的倍数!";
                    return msg;
                }
                return msg;
        }
        #endregion
        /// <summary>
        /// 获取第四行中，为天数的最大列索引
        /// </summary>
        /// <param name="wS"></param>
        /// <returns></returns>
        public static int getMaxColIndexOfThe4thRowOfAR(Worksheet wS) {
            Stack<int> sDate = new Stack<int>();
            sDate.Push(0);
            int aDate = 0;
            int maxColIndex = wS.UsedRange.Columns.Count;
            for (int colIndex = 1; colIndex <= maxColIndex; colIndex++)
            {
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(wS);
                string tempStr = uEHelper.getCellContentByRowAndColIndex(4, colIndex);
                if (string.IsNullOrEmpty(tempStr))
                {
                    return colIndex - 1;
                }
                aDate = int.Parse(tempStr);
                //判断新增的日期是否大于上一个.
                if (aDate <= sDate.Peek())
                {
                    return colIndex - 1;
                }
                sDate.Push(aDate);
                //取其中的最小值。
            }
            return maxColIndex;
        }
        /// <summary>
        /// 判断某行是否都为数字。
        /// </summary>
        /// <param name="wS"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static bool isAllDigit(Worksheet wS, int rowIndex,out int checkedColIndex)
        {
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(wS);
            int maxColIndex = wS.UsedRange.Columns.Count;
            bool flag = false;
            int num = 0;
            for (int colIndex = 1; colIndex <= maxColIndex; colIndex++)
            {
                string tempStr = uEHelper.getCellContentByRowAndColIndex(4, colIndex);
                flag = int.TryParse(tempStr, out num);
                if (!flag) {
                    checkedColIndex = colIndex;
                    return false;
                }
            }
            checkedColIndex = maxColIndex;
            return true;
        }
        /// <summary>
        /// 
        /// </summary>
        public static void clearSheet(Worksheet firstWS) {
            Queue<Range> rangeToDelQueue = new Queue<Range>();
            int rowsMaxCount;
            rowsMaxCount = firstWS.UsedRange.Rows.Count;
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
            //获取最大列
            int maxColIndex = getMaxColIndexOfThe4thRowOfAR(firstWS);
            //判断是否有空行。
            for (int i = 5; i <= rowsMaxCount; i++)
            {
                if (uEHelper.isBlankRangeTheSpecificRow(i, 1, maxColIndex))
                {
                    //只要上一列不是
                    //删除掉此行。
                    //判断上一行中的A列是否为工号。
                    string temp = uEHelper.getSpecificCellValue("A" + (i - 1).ToString());
                    if ("工号:".Equals(temp))
                    {
                        continue;
                    }
                    //获取该行。
                    Range rangeToDel = (Microsoft.Office.Interop.Excel.Range)uEHelper.WS.Rows[i, System.Type.Missing];
                    //不为工号
                    rangeToDelQueue.Enqueue(rangeToDel);
                }
            }
            Range rangeToDelete;
            //开始删除空行。  
            while (rangeToDelQueue.Count > 0)
            {
                rangeToDelete = rangeToDelQueue.Dequeue();
                rangeToDelete.Delete(XlDeleteShiftDirection.xlShiftUp);
            }
        }
      
    }
}
