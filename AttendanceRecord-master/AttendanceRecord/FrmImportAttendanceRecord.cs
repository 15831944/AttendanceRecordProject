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
    /// <summary>
    /// 
    /// </summary>
    public partial class FrmImportAttendanceRecord : Form
    {
        string randomStr = String.Empty;
        public static string _action = "AttendanceRecord";
        string procedureName = "Analyze_AR";
        AttendanceR aR = new AttendanceR();
        /// <summary>
        /// 用于存储过程中的参数。
        /// </summary>
        OracleHelper oH = OracleHelper.getBaseDao();
        string defaultDir = System.Windows.Forms.Application.StartupPath + @"\考勤记录";
        string xlsFilePath = String.Empty;

        MyExcel _1th_my_excel = null;
        MyExcel _2nd_my_excel = null;
        MyExcel _3rd_my_excel = null;
        MyExcel _4th_my_excel = null;
        Worksheet _1th_Sheet = null;
        Worksheet _2nd_Sheet = null;
        Worksheet _3rd_Sheet = null;
        Worksheet _4th_Sheet = null;

        Worksheet uncertainWS = null;
        string _defaultDir = System.Windows.Forms.Application.StartupPath + "\\uncertainRecord";
        
        Range srcRange = null;
        Range destRange = null;
        int currentRow = 0;

        public FrmImportAttendanceRecord()
        {
            InitializeComponent();
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnImportEmpsARInfo_Click(object sender, EventArgs e)
        {
            //判断是否存在Excel进程.
            if (CmdHelper.ifExistsTheProcessByName("EXCEL.EXE")) {
                ShowResult.show(lblResult, "存在未关闭的Office Excel进程,请先关闭!", false);
                return;
            }
            string _uncertainWSPath = _defaultDir + "\\uncertainRecord_" + TimeHelper.getCurrentTimeStr() + ".xls";
            this.dgv.Visible = false;
            this.dgv_same_pinyin_of_name.Visible = true;
            this.dgv_same_name.Visible = false;
            dgv_same_name.DataSource = null;
            dgv_same_pinyin_of_name.DataSource = null;
            dgv.DataSource = null;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            tb.Clear();
            randomStr = TimeHelper.getCurrentTimeStr();
            xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir );
            string dir = DirectoryHelper.getDirOfFile(xlsFilePath);
            if (string.IsNullOrEmpty(dir)) {
                return;
            }
            List<string> xlsFileList = DirectoryHelper.getXlsFileUnderThePrescribedDir(dir);
            List<string> resultList = new List<string>();
            foreach (string xlsFile in xlsFileList) {
                string fileName = DirectoryHelper.getFileNameWithoutSuffix(xlsFile);
                if (!CheckString.CheckARName(fileName)) {
                    continue;
                }
                //格式符合:  3月考勤记录1。
                resultList.Add(xlsFile);
            }
            #region 先判断第四行，是否全为数字。
            int maxColIndex = 0;
            if (!check4thRow(resultList,out maxColIndex)) {
                return;
            }
            #endregion
            if (cbCheckSameNamesButDifferentMachineNo.Checked)
            {
                #region 保存关键信息到后台.
                saveCriticalARInfo(resultList);
                #endregion
                #region  打开4个考勤文件
                for (int i = 1; i <= resultList.Count; i++)
                {
                    switch (i)
                    {
                        case 1:
                            _1th_my_excel = new MyExcel(resultList[0]);
                            _1th_my_excel.open();
                            _1th_Sheet = _1th_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 2:
                            _2nd_my_excel = new MyExcel(resultList[1]);
                            _2nd_my_excel.open();
                            _2nd_Sheet = _2nd_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 3:
                            _3rd_my_excel = new MyExcel(resultList[2]);
                            _3rd_my_excel.open();
                            _3rd_Sheet = _3rd_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                        case 4:
                            _4th_my_excel = new MyExcel(resultList[3]);
                            _4th_my_excel.open();
                            _4th_Sheet = _4th_my_excel.getFirstWorkSheetAfterOpen();
                            break;
                    }
                }
                #endregion
                #region 创建 _uncertain_myExcel;
                MyExcel uncertainRecordExcel = null;
                uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                uncertainRecordExcel.create();
                uncertainRecordExcel.open();
                uncertainWS = uncertainRecordExcel.getFirstWorkSheetAfterOpen();
                //先写，日期行。
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(uncertainWS);
                uEHelper.writeToSpecificRow(1, 1, maxColIndex);
                #endregion
                System.Data.DataTable dt = getSamePinYinButName();
                bool have_same_pinyin_flag = false;
                if (dt != null && dt.Rows.Count > 0)
                {
                    have_same_pinyin_flag = true;
                }
                //*************判断是否拼音相同 开始********************8
                if (have_same_pinyin_flag)
                {
                    this.dgv_same_pinyin_of_name.Visible = true;
                    this.dgv_same_pinyin_of_name.DataSource = dt;
                    DGVHelper.AutoSizeForDGV(dgv_same_pinyin_of_name);
                    ShowResult.show(lblResult, "存在姓名拼音相同的记录!", false);
                    this.lblPrompt.Visible = false;
                    timerRestoreTheLblResult.Enabled = true;
                    btnSwitch.Text = "查看拼音相同的记录";
                    btnSwitch_Click(sender, e);
                    #region 写记录到不确定文档中.
                    int theRowIndex = 0;
                    int Attendance_Machine_No = 0;
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        theRowIndex = int.Parse(dt.Rows[i]["行号"].ToString());
                        Attendance_Machine_No = int.Parse(dt.Rows[i]["卡机编号"].ToString());
                        switch (Attendance_Machine_No)
                        {
                            case 1:
                                //获取源区域
                                srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;

                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                //保存一下。
                                break;
                            case 2:
                                srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
   
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 3:
                                srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 4:
                                srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                        }
                    }
                    //显示该文档
                    CmdHelper.runCmd(_uncertainWSPath);
                    #endregion
                    string promptStr = string.Format(@"存在姓名拼音相同的记录。{0}
                                                                              继续(OK), 取消导入(Cancel)。", "\r\n");
                    if (DialogResult.Cancel.Equals(MessageBox.Show(promptStr, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                    {
                        uncertainWS.UsedRange.ColumnWidth = 3.75M;
                        uncertainRecordExcel.saveWithoutAutoFit();
                        uncertainRecordExcel.close();
                        closeThe4ARExcels();
                        return;
                    }
                }
                //closeThe4ARExcels();
                //*************判断是否拼音相同  结束*****************88
                //1.h
                dt = getSameNameInfo();
                string prompt = string.Empty;
                if (dt.Rows.Count != 0)
                {
                    btnSwitch.Text = "查看同名记录";
                    btnSwitch_Click(sender, e);

                    int theRowIndex = 0;
                    int Attendance_Machine_No = 0;
                    #region 同名记录书写结束.
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {
                        theRowIndex = int.Parse(dt.Rows[i]["行号"].ToString());
                        Attendance_Machine_No = int.Parse(dt.Rows[i]["卡机编号"].ToString());
                        switch (Attendance_Machine_No)
                        {
                            case 1:
                                //获取源区域
                                srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                //保存一下。
                                break;
                            case 2:
                                srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 3:
                                srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                            case 4:
                                srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndex]];
                                srcRange.Cells.Copy(Type.Missing);
                                //向目标复制。
                                //或取目标单元格。
                                currentRow = uncertainWS.UsedRange.Rows.Count;
                                destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndex]];
                                //destRange.Select();
                                uncertainWS.Paste(destRange, false);
                                break;
                        }
                    }
                    #endregion
                    prompt = string.Format(@"存在同名记录，将合并;{0}
                                                                              继续(OK), 取消导入(Cancel)。", "\r\n");
                    if (DialogResult.Cancel.Equals(MessageBox.Show(prompt, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                    {
                        uEHelper.setAllColumnsWidth(3.75M);
                        uncertainRecordExcel.saveWithoutAutoFit();
                        uncertainRecordExcel.close();
                        closeThe4ARExcels();
                        return;
                    }
                }
                //关闭不确定文档。
                uEHelper.setAllColumnsWidth(3.75M);
                uncertainRecordExcel.saveWithoutAutoFit();
                uncertainRecordExcel.close();
                closeThe4ARExcels();
            }
            resultList.Sort();
            //判断该考勤机中是否已经存在，某月的记录
            foreach (string xlsFilePath in resultList)
            {
                string fileName = DirectoryHelper.getFileNameWithoutSuffix(xlsFilePath);
                int attendanceMachineFlag = int.Parse(fileName.Substring(fileName.Length - 1, 1));
                //打开文档获取考勤机，所记录的日期。
                MyExcel myExcel = new MyExcel(xlsFilePath);
                myExcel.open();
                Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstSheet);
                string year_and_month_str = uEHelper.getCellContentByRowAndColIndex(3, 3);
                year_and_month_str = year_and_month_str.Substring(0, 7);
                myExcel.close();
                System.Data.DataTable dtSubmitInfo = getSubmitInfoOfTheSpecificeMachineAndYearAndMonth(attendanceMachineFlag, year_and_month_str);
                if (!"0".Equals(dtSubmitInfo.Rows[0]["nums_of_staffs"].ToString())) {
                    string prompt = string.Format(@"您已经于: {1},{0}提交了{2} 考勤机{3},{0}共计{4}个用户的纪录{0}    覆盖(OK),退出(Cancel)?","\r\n",dtSubmitInfo.Rows[0]["latest_record_time"].ToString(),year_and_month_str,attendanceMachineFlag,dtSubmitInfo.Rows[0]["nums_of_staffs"].ToString());
                    if (DialogResult.Cancel.Equals(MessageBox.Show(prompt, "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Question))){
                        return;
                    }
                    delTheInfoOfTheSpecificeMachineAndYearAndMonth(attendanceMachineFlag, year_and_month_str);
                }
            }
            this.dgv.DataSource = null;
            //this.dgv.Columns.Clear();
            lblPrompt.Visible = false;
            lblPrompt.Text = "";
            pb.Value = 0;
            pb.Maximum = 0;
            pb.Visible = false;
            foreach (string xlsFilePath in resultList) {
                tb.Text = xlsFilePath;
                lblResult.Visible = false;
                MSG msg = AttendanceRHelper.ImportAttendanceRecordToDB(xlsFilePath, randomStr, lblPrompt,  pb,lblResult);
                //导入完成后进行保存，保存该文件至prepared目录中
                pb.Visible = false;
                lblPrompt.Visible = false;
                ShowResult.show(lblResult, msg.Msg, msg.Flag);
                timerRestoreTheLblResult.Enabled = true;
                if (!msg.Flag) return;
                //saveTheExcel(xlsFilePath);
            }
            dgv_same_name.Visible = false;
            dgv_same_pinyin_of_name.Visible = false;
            //加载导入的数据。
            this.dgv.DataSource = null;
            this.dgv.DataSource = aR.getAR(randomStr);
            this.dgv.Visible = true;
            DGVHelper.AutoSizeForDGV(dgv);
            tb.Clear();
        }
        private void closeThe4ARExcels()
        {
            //检查结束.
            #region 关闭4个考勤文件
            if (_1th_my_excel != null)
            {
                _1th_my_excel.close();
            }
            if (_2nd_my_excel != null)
            {
                _2nd_my_excel.close();
            }
            if (_3rd_my_excel != null)
            {
                _3rd_my_excel.close();
            }
            if (_4th_my_excel != null)
            {
                _4th_my_excel.close();
            }
            #endregion
        }
        private System.Data.DataTable getSameNameInfo()
        {
            string sqlStr = string.Format(@"select distinct 
                                                            AR_Temp.Job_Number AS ""工号"",
                                                            AR_Temp.name AS ""姓名"",
                                                            AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                            AR_Temp.Row_Index AS ""行号"",
                                                            AR_Temp.Record_Time  AS ""记录时间""
                                            from AR_Temp, (
                                                           select job_number,
                                                                    name,
                                                                    attendance_machine_flag,
                                                                    row_index,
                                                                    record_time
                                                           from AR_Temp
                                             ) AR_T
                                            WHERE AR_Temp.name = AR_T.Name
                                            AND(
                                                  (AR_Temp.Attendance_Machine_Flag = AR_T.attendance_machine_flag
                                                  AND AR_Temp.Job_Number != AR_T.job_number)
                                                  OR(
                                                     AR_Temp.Attendance_Machine_Flag != AR_T.attendance_machine_flag
                                                  )
                                            )
                                            order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            this.dgv_same_name.DataSource = dt;
            DGVHelper.AutoSizeForDGV(dgv_same_name);
            return dt;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private MSG checkExistSameStaffsButDifferentAttendanceMachine(List<String> excelPathList)
        {
            MSG msg = new MSG();
            List<SimpleARInfo> simpleARInfoList = new List<SimpleARInfo>();
            foreach (string excelPath in excelPathList) {
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                //删除  时间为空 的行。
                AttendanceRHelper.clearSheet(firstWS);
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
               
                SimpleARInfo simpleARInfo = null;
                string fileNameWithout = DirectoryHelper.getFileNameWithoutSuffix(excelPath);

                int rowMaxIndex = firstWS.UsedRange.Rows.Count;
                for (int rowIndex = 5;rowIndex<=rowMaxIndex;rowIndex++) {
                    if (0 == rowIndex % 2) continue;
                    //姓名 存于第11列。
                    string name = uEHelper.getCellContentByRowAndColIndex(rowIndex, 11);
                    simpleARInfo = new SimpleARInfo();
                    simpleARInfo.AttendanceMachineFlag = int.Parse(fileNameWithout.Substring(fileNameWithout.Length - 1, 1));
                    simpleARInfo.Name = name;
                    simpleARInfo.RowIndex = rowIndex;
                    SimpleARInfo sARInfo = simpleARInfoList.Find(x =>x.Name == simpleARInfo.Name);
                    if (sARInfo != null) {
                        //存在重复的员工.
                        msg.Msg = string.Format(@"{0} 与 {1} 同名;同一人请汇总为一行,不同人,请姓名相异。",sARInfo.toString(),simpleARInfo.toString());
                        myExcel.close();
                        return msg;
                    }
                    simpleARInfoList.Add(simpleARInfo);
                }
                myExcel.save();
                myExcel.close();
            }
            msg.Flag = true;
            msg.Msg = "未发现同名用户";
            return msg;
        }
        /// <summary>
        /// 检查第四行
        /// </summary>
        /// <returns></returns>
        private bool check4thRow(List<String> excelPathList,out int maxColIndex) {
            maxColIndex = 0;
            //先清除所有记录。
            AR_Temp.deleteTheARTemp();
            foreach (string excelPath in excelPathList)
            {
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                string fileNameWithoutSuffix = DirectoryHelper.getFileNameWithoutSuffix(excelPath);
                int checkedColIndex = 0;
                if (!AttendanceRHelper.isAllDigit(firstWS, 4,out checkedColIndex))
                {
                    myExcel.close();
                    lblPrompt.Visible = false;
                    ShowResult.show(lblResult, fileNameWithoutSuffix + ": 第4行"+checkedColIndex.ToString()+"列非数字;   导入取消。", false);
                    timerRestoreTheLblResult.Start();
                    return false;
                }
                if (maxColIndex == 0)
                {
                    Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
                    maxColIndex = uEHelper.getMaxColIndexBeforeBlankCellInSepcificRow(4);
                }
                myExcel.close();
            }
            return true;
        }
        /// <summary>
        /// 保存关键的 考勤信息.
        /// </summary>
        private void saveCriticalARInfo(List<String> excelPathList) {
            //先清除所有记录。
            AR_Temp.deleteTheARTemp();
            foreach (string excelPath in excelPathList)
            {
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                //删除  时间后立即为空的行。
                AttendanceRHelper.clearSheet(firstWS);
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
                string fileNameWithoutSuffix = DirectoryHelper.getFileNameWithoutSuffix(excelPath);
                //先获取第4行的最大行列数目。
                int rowMaxIndex = firstWS.UsedRange.Rows.Count;
                pb.Value = 0;
                pb.Maximum = rowMaxIndex - 4;
                pb.Visible = true;
                lblPrompt.Visible = true;
                lblPrompt.Text = fileNameWithoutSuffix + ": 基本信息采集中...";
                for (int rowIndex = 5; rowIndex <= rowMaxIndex; rowIndex++)
                {
                    if (0 == rowIndex % 2)
                    {
                        pb.Value++;
                        continue;
                    }
                    //姓名 存于第11列。
                    string name = uEHelper.getCellContentByRowAndColIndex(rowIndex, 11);
                    AR_Temp ar_Temp = new AR_Temp();
                    ar_Temp.Attendance_machine_flag = int.Parse(fileNameWithoutSuffix.Substring(fileNameWithoutSuffix.Length - 1, 1));
                    ar_Temp.Row_Index = rowIndex;
                    ar_Temp.Job_number = uEHelper.getCellContentByRowAndColIndex(rowIndex, 3);
                    ar_Temp.Name = name;
                    ar_Temp.saveRecord();
                    pb.Value++;
                }
                lblPrompt.Visible = false;
                pb.Visible = false;
                myExcel.save();
                myExcel.close();
            }
        }
        /// <summary>
        /// 获取拼音相同但名字不同的用户信息.
        /// </summary>
        /// <returns></returns>
        private System.Data.DataTable  getSamePinYinButName()
        {
            string sqlStr = string.Format(@"select distinct 
                                                                AR_Temp.Job_Number AS ""工号"",
                                                                AR_Temp.name AS ""姓名"",
                                                                AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                                AR_Temp.Row_Index AS ""行号"",
                                                                AR_Temp.Record_Time  AS ""记录时间""
                                                from AR_Temp, (
                                                              select distinct name
                                                              from AR_Temp
                                                ) AR_T
                                                WHERE AR_Temp.name ! = AR_T.Name
                                                AND gethzpy.GetHzFullPY(AR_Temp.name) = gethzpy.GetHzFullPY(AR_T.name)
                                                order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        /// <summary>
        /// 删除指定月份的 某个考勤机对应的纪录
        /// </summary>
        /// <param name="attendanceMachineFlag"></param>
        /// <param name="year_and_month_str"></param>
        private void delTheInfoOfTheSpecificeMachineAndYearAndMonth(int attendanceMachineFlag, string year_and_month_str)
        {
            string sqlStr = string.Format(@"delete 
                                            from Attendance_record 
                                            where substr(job_number,1,1) = '{0}'
                                            and trunc(start_date,'MM') = to_date('{1}','yyyy-MM')",
                                            attendanceMachineFlag,
                                            year_and_month_str);
            OracleDaoHelper.executeSQL(sqlStr);
        }
        /// <summary>
        /// 获取该考勤机在后台已经存在的信息.
        /// </summary>
        /// <param name="attendanceMachineFlag"></param>
        /// <param name="year_and_month_str"></param>
        /// <returns></returns>
        private System.Data.DataTable  getSubmitInfoOfTheSpecificeMachineAndYearAndMonth(int attendanceMachineFlag, string year_and_month_str)
        {
            string sqlStr = string.Format(@"Select count(distinct job_number) as nums_of_staffs,
                                                    max(to_char(record_time,'yyyy-MM-dd HH24:MI:SS')) as latest_record_time
                                             from Attendance_Record
                                             where substr(job_number,1,1) = '{0}'
                                             and trunc(fingerprint_date,'MM') = to_date('{1}','yyyy-MM')",
                                             attendanceMachineFlag,
                                             year_and_month_str);
            return OracleDaoHelper.getDTBySql(sqlStr);
        }
        /// <summary>
        /// 保存一份文件到Prepared中。
        /// </summary>
        /// <param name="xlsFilePath"></param>
        private void saveTheExcel(string xlsFilePath)
        {
            //获取该记录表所对应的月份。
            MyExcel srcExceFile = new MyExcel(xlsFilePath);
            srcExceFile.openWithoutAlerts();
            Worksheet firstSheet_Src = srcExceFile.getFirstWorkSheetAfterOpen();
            string year_and_month_src_str = new Usual_Excel_Helper(firstSheet_Src).getCellContentByRowAndColIndex(3, 3).Substring(0, 7);
            string fileNameWithoutSuffix = DirectoryHelper.getFileNameWithoutSuffix(xlsFilePath);
            string attendanceMachineFlag = fileNameWithoutSuffix.Substring(fileNameWithoutSuffix.Length - 1);
            year_and_month_src_str += string.Format(@"_{0}", attendanceMachineFlag);
            srcExceFile.close();
            //year_and_month_src_str  即为文件名

            //string fileName = DirectoryHelper.getFileName(xlsFilePath);
            //1.复制该excel 到 Prepared中。
            Tools.CmdHelper.copyFileToDestDir(xlsFilePath, System.Windows.Forms.Application.StartupPath + "\\prepared\\"+ year_and_month_src_str+".xls");
            string destFilePath = System.Windows.Forms.Application.StartupPath + "\\prepared\\" + year_and_month_src_str + ".xls";
            //打开该文件
            MyExcel myExcel = new MyExcel(destFilePath);
            myExcel.openWithoutAlerts();
            //新增一个sheet.
            Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
            Usual_Excel_Helper uHelper = new Usual_Excel_Helper(firstSheet);
            //获取月份.
            string C3ContentStr = uHelper.getCellContentByRowAndColIndex(3, 3);
            string year_and_month_str = C3ContentStr.Substring(0, 7);
            Worksheet theLastestExcel = myExcel.AddSheetToLastIndex(year_and_month_str);
            myExcel.copyRangeFromOneToAnotherSheet(firstSheet, theLastestExcel);
            int sheetsCount = myExcel.getCountsOfAllSheet();
            //删掉之前的表格,保留最后一个。
            for (int i = 1; i <= sheetsCount-1; i++) {
                myExcel.delTheSheet(1);
            }
            //一定要保存否则，无效。
            myExcel.saveAndColumnsAutoFit();
            //关闭该文件
            myExcel.close();
        }
        private System.Data.DataTable getAR(string random_Str) {
            String sqlStr = String.Format(@"SELECT 
                                     start_date; 
                                      end_date; 
                                      tabulation_time; 
                                      fingerprint_date; 
                                      dept;
                                      job_number; 
                                      name; 
                                      sheet_name; 
                                      fpt_first_time; 
                                      fpt_last_time; 
                                      seq; 
                                      reord_time; 
                                      not_finger_print; 
                                      delay_time; 
                                      come_late;
                                      leave_early; 
                                      dinner_subsidy; 
                                      random_str 
                              FROM Attendance_Record
                              WHERE Random_Str= '{0}'
                              OrDER bY 
                                        FINGERPRINT_DATE DESC;
                                         Dept DESC;
                                        JOB_NUMBER DESC
                                         ", random_Str);
            System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
            return dt;
        } 
        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            lblResult.Visible = false;
        }
        /// <summary>
        /// 转换列名
        /// </summary>
        private void ConvertColumnsName(System.Data.DataTable dt) {
            dt.Columns["JOB_NUMBER"].ColumnName = "工号";
            dt.Columns["NAME"].ColumnName = "姓名";
            dt.Columns["Department"].ColumnName = "部门";
            dt.Columns["POSITION"].ColumnName = "职位";
            dt.Columns["Role"].ColumnName = "角色";
            dt.Columns["Contact_Way"].ColumnName = "联系方法";
            dt.Columns["UPDATE_Time"].ColumnName = "更新时间";
        }
        /// <summary>
        /// 分析考勤数据。
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            int V_FLAG = 0;
            //分析刚刚导入的数据。
            OracleParameter parma_RandomStr = new OracleParameter("v_RandomStr", OracleDbType.NVarchar2, ParameterDirection.Input);
            OracleParameter param_V_FLAG = new OracleParameter("v_FLAG", OracleDbType.Int32, ParameterDirection.Output);
            parma_RandomStr.Value = randomStr;
            param_V_FLAG.Value = V_FLAG;
            OracleParameter[] parameters = new OracleParameter[2] { parma_RandomStr, param_V_FLAG };
            int j = oH.ExecuteNonQuery(procedureName, parameters);
            V_FLAG = Int32.Parse(parameters[1].Value.ToString());
            if (V_FLAG == 1)
            {
                //获取导入的信息
                ShowResult.show(lblResult, "分析结束", true);
                timerRestoreTheLblResult.Enabled = true;
            }
            else
            {
                /*
                //未能导入，展示查询结果
                DataTable dt1 = OracleDaoHelper.getDTBySql(String.Format(@"SELECT * FROM MESSAGE WHERE IPADDR = '{0}' AND UPPER(Subject) = UPPER('{1}') AND Random_Str = '{2}' ORDER BY OPERATE_TIME DESC;PROMPT DESC"; Program.IPADDR; subject; _staffTemp.Random_Str));
                this.dgv.DataSource = null;
                Messages.convertColumnsName(dt1);
                dgv.DataSource = dt1;
                DGVHelper.AutoSizeForDGV(dgv);
                ShowResult.prompt(lblResult; "存在同名员工，确定导入，请点击\"确定导入\"按钮!"; false);
                timerRestoreTheLblResult.Enabled = true;
                */
            }
        }
        private void btnSwitch_Click(object sender, EventArgs e)
        {
            if ("查看同名记录".Equals(btnSwitch.Text.Trim()))
            {
                btnSwitch.Text = "查看拼音相同的记录";
                dgv_same_name.Visible = true;
                dgv_same_pinyin_of_name.Visible = false;
                string sqlStr = string.Format(@"select distinct 
                                                            AR_Temp.Job_Number AS ""工号"",
                                                            AR_Temp.name AS ""姓名"",
                                                            AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                            AR_Temp.Row_Index AS ""行号"",
                                                            AR_Temp.Record_Time  AS ""记录时间""
                                            from AR_Temp, (
                                                           select job_number,
                                                                    name,
                                                                    attendance_machine_flag,
                                                                    row_index,
                                                                    record_time
                                                           from AR_Temp
                                             ) AR_T
                                            WHERE AR_Temp.name = AR_T.Name
                                            AND(
                                                  (AR_Temp.Attendance_Machine_Flag = AR_T.attendance_machine_flag
                                                  AND AR_Temp.Job_Number != AR_T.job_number)
                                                  OR(
                                                     AR_Temp.Attendance_Machine_Flag != AR_T.attendance_machine_flag
                                                  )
                                            )
                                            order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
                System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
                this.dgv_same_name.DataSource = dt;
                DGVHelper.AutoSizeForDGV(dgv_same_name);
            }
            else
            {
                btnSwitch.Text = "查看同名记录";
                string sqlStr = string.Format(@"select distinct 
                                                                AR_Temp.Job_Number AS ""工号"",
                                                                AR_Temp.name AS ""姓名"",
                                                                AR_Temp.Attendance_Machine_Flag AS ""卡机编号"",
                                                                AR_Temp.Row_Index AS ""行号"",
                                                                AR_Temp.Record_Time  AS ""记录时间""
                                                from AR_Temp, (
                                                              select distinct name
                                                              from AR_Temp
                                                ) AR_T
                                                WHERE AR_Temp.name ! = AR_T.Name
                                                AND gethzpy.GetHzFullPY(AR_Temp.name) = gethzpy.GetHzFullPY(AR_T.name)
                                                order by NLSSORT(""姓名"", 'NLS_SORT= SCHINESE_PINYIN_M')");
                System.Data.DataTable dt = OracleDaoHelper.getDTBySql(sqlStr);
                this.dgv_same_pinyin_of_name.DataSource = dt;
                dgv_same_pinyin_of_name.Visible = true;
                DGVHelper.AutoSizeForDGV(dgv_same_pinyin_of_name);
                dgv_same_name.Visible = false;
            }
        }
    }
}

