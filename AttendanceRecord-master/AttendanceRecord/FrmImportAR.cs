using AttendanceRecord.Entities;
using AttendanceRecord.Helper;
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
using System.IO;
using System.Threading;

namespace AttendanceRecord
{
    public partial class FrmImportAR : Form
    {
        string randomStr = String.Empty;
        public static string _action = "importAR";
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
        string _uncertainWSPath = string.Empty;
        Range srcRange = null;
        Range destRange = null;
        int currentRow = 0;
        string currentXlsFilePath = string.Empty;

        BackgroundWorker readDataFromExcelBGWorker = new BackgroundWorker();
        BackgroundWorker checkNameBGWorker = new BackgroundWorker();
        List<String> xlsFilePathList = null;
        //当前正在处理的 excel 路径。
        //导入时后台程序汇报的消息。
        MSG msg = new MSG();
        int maxColIndexOfCheckedNameOfExcel = 0;
        Thread threadG;//声明线程
        delegate void changetext(System.Data.DataTable result);
        //创建一个委托
        public delegate void UpdateD(System.Data.DataTable result);
        public UpdateD updated;
        private void dgv_Load()
        {
            updated = new UpdateD(UpdateDataTable);
        }
        public void Getdgv()
        {
            System.Data.DataTable result = aR.getAR(randomStr);
            //  CalcFinished(result);
            Invoke(new changetext(Changetext), new object[] { result });
        }
        public void Changetext(System.Data.DataTable result)
        {
            dgv.DataSource = result;
            DGVHelper.AutoSizeForDGV(dgv);
        }
        public void UpdateDataTable(System.Data.DataTable result)
        {
            dgv.DataSource = result;
        }
        public void CalcFinished(System.Data.DataTable result)
        {
            if (this.dgv.InvokeRequired)
            {
                while (!this.dgv.IsHandleCreated)
                {
                    if (this.dgv.Disposing || this.dgv.IsDisposed)
                    {
                        return;
                    }
                }
                changetext c = new changetext(CalcFinished);
                this.dgv.Invoke(c, new object[] { result });

                //   OutPutResult(_grpGetData.Text, _btnGetMeasureData.Text, result);
            }
            else
            {
                dgv.DataSource = result;
                //  dataGridView1.Columns["序号"].DisplayIndex = 0;
            }
        }
        public FrmImportAR()
        {
            InitializeComponent();
            readDataFromExcelBGWorker.DoWork += DoWork_Handler;
            readDataFromExcelBGWorker.ProgressChanged += ProcessChanged_Handler;
            readDataFromExcelBGWorker.RunWorkerCompleted += RunWorkerCompleted_Handler;
            readDataFromExcelBGWorker.WorkerReportsProgress = true;
            checkNameBGWorker.DoWork += CheckName_DoWork_Handler;
            checkNameBGWorker.ProgressChanged += ProcessChanged_Handler;
            checkNameBGWorker.RunWorkerCompleted += CheckName_RunWorkerCompleted_Handler;
            checkNameBGWorker.WorkerReportsProgress = true;
        }
        private void CheckName_RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs e)
        {
            lblPrompt.Visible = false;
            pb.Visible = false;
            doNextAfterCheckName();
        }
        private void CheckName_DoWork_Handler(object sender, DoWorkEventArgs e)
        {
            saveCriticalARInfo(xlsFilePathList);
        }
        private void RunWorkerCompleted_Handler(object sender, RunWorkerCompletedEventArgs e)
        {
            V_Summary_OF_AR.generateAttendanceRecordBriefly(AttendanceRecordDetail._start_date.Substring(0, 7));
            //加载导入的数据。
            ThreadStart threaddgv = new ThreadStart(Getdgv);
            threadG = new Thread(threaddgv);
            //  threadG.IsBackground = true;
            threadG.Start();
            this.dgv.Visible = true;
            tb.Clear();
            this.btnImportEmpsInfo.Enabled = true;
        }
        private void ProcessChanged_Handler(object sender, ProgressChangedEventArgs e)
        {
            if (e.UserState.ToString().Contains("lblResult.Visible"))
            {
                if (0 == e.ProgressPercentage)
                {
                    lblResult.Visible = false;
                }
                else {
                    lblResult.Visible = true;
                }
            }
            if ("pb.Maximum".Equals(e.UserState.ToString())){
                pb.Maximum = e.ProgressPercentage;
                pb.Visible = true;
                lblResult.Visible = false;
            }
            if ("pb.Value".Equals(e.UserState.ToString())) {
                pb.Value = e.ProgressPercentage;
                pb.Visible = true;
                lblResult.Visible = false;
            }
            if (e.UserState.ToString().Contains("lblPrompt.Text")){
                lblPrompt.Text = e.UserState.ToString().Split('=')[1];
                if (string.IsNullOrEmpty(lblPrompt.Text.Trim()))
                {
                    lblPrompt.Visible = false;
                }
                else
                {
                    lblPrompt.Visible = true;
                }
            }
            if (e.UserState.ToString().Contains("lblResult.Text"))
            {
                lblResult.Text = e.UserState.ToString().Split('=')[1];
                if (string.IsNullOrEmpty(lblResult.Text.Trim()))
                {
                    lblResult.Visible = false;
                }
                else {
                    lblResult.Visible = true;
                }
            }
            if (e.UserState.ToString().Contains("tb.Text")) {
                tb.Text = e.UserState.ToString().Split('=')[1];
            }
            //格式： msg.msg = "":"true";
            if (e.UserState.ToString().Contains("msg=")) {
                msg.Msg = e.UserState.ToString().Split('=')[1];
                msg.Flag = e.ProgressPercentage == 1 ? true : false;
                ShowResult.show(lblPrompt, pb, lblResult, msg.Msg, msg.Flag);
            }
        }
        private void DoWork_Handler(object sender, DoWorkEventArgs e)
        {
            AttendanceRHelper.affectedCount = 0;
            foreach (string xlsFilePath in xlsFilePathList)
            {
                currentXlsFilePath = xlsFilePath;
                //tb.Text = xlsFilePath;
                readDataFromExcelBGWorker.ReportProgress(0,"tb.Text=" + xlsFilePath);
                //lblResult.Visible = false;
                //开启后台执行进程。
                MSG msg = AttendanceRHelper.ImportAttendanceRecordToDB(currentXlsFilePath, randomStr, readDataFromExcelBGWorker);
                //导入完成后进行保存，保存该文件至prepared目录中
                //pb.Visible = false;
                //lblPrompt.Visible = false;
                //
                readDataFromExcelBGWorker.ReportProgress(msg.Flag?1:0, string.Format(@"msg={0}",msg.Msg));
                //timerRestoreTheLblResult.Enabled = true;
                if (!msg.Flag)
                {
                    return;
                }
            }
        }
        private void btnImportEmpsInfo_Click(object sender, EventArgs e)
        {
            btnViewTheUncertaiRecordInExcel.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            lblResult.Visible = false;
            //判断是否存在Excel进程.
            if (CmdHelper.ifExistsTheProcessByName("EXCEL"))
            {
                FrmPrompt frmPrompt = new FrmPrompt();
                frmPrompt.ShowDialog();
            }
            _uncertainWSPath = _defaultDir + "\\uncertainRecord_" + TimeHelper.getCurrentTimeStr() + ".xls";
            dgv.DataSource = null;
            lblResult.Visible = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            tb.Clear();
            randomStr = TimeHelper.getCurrentTimeStr();
            xlsFilePath = FileNameDialog.getSelectedFilePathWithDefaultDir("请选择考勤记录：", "*.xls|*.xls", defaultDir);
            string dir = DirectoryHelper.getDirOfFile(xlsFilePath);
            if (string.IsNullOrEmpty(dir))
            {
                return;
            }
            List<string> xlsFileList = DirectoryHelper.getXlsFileUnderThePrescribedDir(dir);
            xlsFilePathList = new List<string>();
            foreach (string xlsFile in xlsFileList)
            {
                string fileName = DirectoryHelper.getFileNameWithoutSuffix(xlsFile);
                if (!CheckString.CheckARName(fileName))
                {
                    continue;
                }
                //格式符合:  3月考勤记录1。
                xlsFilePathList.Add(xlsFile);
            }
            #region 先判断第四行，是否全为数字。
            if (!check4thRow(xlsFilePathList, out maxColIndexOfCheckedNameOfExcel))
            {
                return;
            }
            #endregion
            #region 保存关键信息到后台.
            checkNameBGWorker.RunWorkerAsync();
            #endregion
            //开启后台工作者
        }
        private void doNextAfterCheckName() {
            #region  打开4个考勤文件
            for (int i = 1; i <= xlsFilePathList.Count; i++)
            {
                switch (i)
                {
                    case 1:
                        _1th_my_excel = new MyExcel(xlsFilePathList[0]);
                        _1th_my_excel.open();
                        _1th_Sheet = _1th_my_excel.getFirstWorkSheetAfterOpen();
                        break;
                    case 2:
                        _2nd_my_excel = new MyExcel(xlsFilePathList[1]);
                        _2nd_my_excel.open();
                        _2nd_Sheet = _2nd_my_excel.getFirstWorkSheetAfterOpen();
                        break;
                    case 3:
                        _3rd_my_excel = new MyExcel(xlsFilePathList[2]);
                        _3rd_my_excel.open();
                        _3rd_Sheet = _3rd_my_excel.getFirstWorkSheetAfterOpen();
                        break;
                    case 4:
                        _4th_my_excel = new MyExcel(xlsFilePathList[3]);
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
            uncertainRecordExcel.openWithoutAlerts();
            uncertainWS = uncertainRecordExcel.getFirstWorkSheetAfterOpen();
            //先写，日期行。
            Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(uncertainWS);
            uEHelper.writeToSpecificRow(1, 1, maxColIndexOfCheckedNameOfExcel);
            #endregion
            System.Data.DataTable dt = getSamePinYinButName();
            int amountOfGroupOfSamePinYinButName = getAmountOfGroupOfSamePinYinButName();
            bool have_same_pinyin_flag = false;
            if (dt != null && dt.Rows.Count > 0)
            {
                have_same_pinyin_flag = true;
            }
            //*************判断是否拼音相同 开始********************8
            if (have_same_pinyin_flag)
            {
                ShowResult.show(lblResult, "存在姓名拼音相同的记录!", false);
                this.lblPrompt.Visible = false;
                timerRestoreTheLblResult.Enabled = true;
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
                            //替换源文件的工号为  工号位于第三列
                            _1th_Sheet.Cells[theRowIndex, 3] = "'111111111" + ((Range)(_1th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;

                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            //保存一下。
                            break;
                        case 2:
                            _2nd_Sheet.Cells[theRowIndex, 3] = "'222222222" + ((Range)(_2nd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;

                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                        case 3:
                            _3rd_Sheet.Cells[theRowIndex, 3] = "'333333333" + ((Range)(_3rd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                        case 4:
                            _4th_Sheet.Cells[theRowIndex, 3] = "'444444444" + ((Range)(_4th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                    }
                }
                //设置列宽
                uncertainWS.UsedRange.ColumnWidth = 3.75;
                //显示该文档对应的图片
                #endregion
                string promptStr = string.Format(@" 存在姓名拼音相同但书写不同的记录：{1}组;{0}
确定: 将视为不同员工;   取消: 取消本次导入;", "\r\n", amountOfGroupOfSamePinYinButName);
                if (DialogResult.Cancel.Equals(MessageBox.Show(promptStr, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                {
                    closeThe4ARExcels();
                    uncertainWS.UsedRange.ColumnWidth = 3.75M;
                    uncertainRecordExcel.saveWithoutAutoFit();

                    uncertainRecordExcel.close();
                    //显示该文档。
                    uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                    uncertainRecordExcel.open(true);
                    btnViewTheUncertaiRecordInExcel.Enabled = true;
                    return;
                }
                if (!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true;
            }

            //*************判断是否拼音相同  结束*****************88
            //1.h
            dt = getSameNameInfo();
            //获取汉字相同的组的数目。
            int amountOfGroupOfSameName = getAmountOfGroupOfSameName();
            string prompt = string.Empty;
            if (dt.Rows.Count != 0)
            {
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
                            _1th_Sheet.Cells[theRowIndex, 3] = "'111111111" + ((Range)(_1th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            //获取源区域
                            srcRange = _1th_Sheet.Range[_1th_Sheet.Cells[theRowIndex, 1], _1th_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            //保存一下。
                            break;
                        case 2:
                            _2nd_Sheet.Cells[theRowIndex, 3] = "'222222222" + ((Range)(_2nd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _2nd_Sheet.Range[_2nd_Sheet.Cells[theRowIndex, 1], _2nd_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                        case 3:
                            _3rd_Sheet.Cells[theRowIndex, 3] = "'333333333" + ((Range)(_3rd_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _3rd_Sheet.Range[_3rd_Sheet.Cells[theRowIndex, 1], _3rd_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                        case 4:
                            _4th_Sheet.Cells[theRowIndex, 3] = "'444444444" + ((Range)(_4th_Sheet.Cells[theRowIndex, 3])).Text.ToString().PadLeft(3, '0');
                            srcRange = _4th_Sheet.Range[_4th_Sheet.Cells[theRowIndex, 1], _4th_Sheet.Cells[theRowIndex + 1, maxColIndexOfCheckedNameOfExcel]];
                            srcRange.Cells.Copy(Type.Missing);
                            //向目标复制。
                            //或取目标单元格。
                            currentRow = uncertainWS.UsedRange.Rows.Count;
                            destRange = uncertainWS.Range[uncertainWS.Cells[currentRow + 1, 1], uncertainWS.Cells[currentRow + 2, maxColIndexOfCheckedNameOfExcel]];
                            //destRange.Select();
                            uncertainWS.Paste(destRange, false);
                            break;
                    }
                }
                #endregion
                prompt = string.Format(@"  存在同名的记录：{1}组;{0}
确定: 将视为同一员工;   取消: 取消本次导入;", "\r\n", amountOfGroupOfSameName);
                if (DialogResult.Cancel.Equals(MessageBox.Show(prompt, "提示：", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)))
                {
                    uEHelper.setAllColumnsWidth(3.75M);
                    uncertainRecordExcel.saveWithoutAutoFit();
                    uncertainRecordExcel.close();
                    //显示该文档。
                    uncertainRecordExcel = new MyExcel(_uncertainWSPath);
                    uncertainRecordExcel.open(true);
                    closeThe4ARExcels();
                    if (!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true;
                    return;
                }
                if (!btnViewTheUncertaiRecordInExcel.Enabled) btnViewTheUncertaiRecordInExcel.Enabled = true;
            }
            //关闭不确定文档。
            uEHelper.setAllColumnsWidth(3.75M);
            uncertainRecordExcel.saveWithoutAutoFit();
            uncertainRecordExcel.close();
            closeThe4ARExcels();

            xlsFilePathList.Sort();

            //直接删除
            MyExcel myExcel = new MyExcel(xlsFilePath);
            myExcel.open();
            Worksheet firstSheet = myExcel.getFirstWorkSheetAfterOpen();
            Usual_Excel_Helper uEHelperTemp = new Usual_Excel_Helper(firstSheet);
            string year_and_month_str = uEHelperTemp.getCellContentByRowAndColIndex(3, 3);
            year_and_month_str = year_and_month_str.Substring(0, 7);
            myExcel.close();
            delARDetailInfoByYearAndMonth(year_and_month_str);
            //删除完毕。
            this.dgv.DataSource = null;
            //this.dgv.Columns.Clear();
            lblPrompt.Visible = false;
            lblPrompt.Text = "";
            pb.Value = 0;
            pb.Maximum = 0;
            pb.Visible = false;
            this.btnImportEmpsInfo.Enabled = false;
            readDataFromExcelBGWorker.RunWorkerAsync();
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
        private bool check4thRow(List<String> excelPathList, out int maxColIndex)
        {
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
                if (!AttendanceRHelper.isAllDigit(firstWS, 4, out checkedColIndex))
                {
                    myExcel.close();
                    lblPrompt.Visible = false;
                    ShowResult.show(lblResult, fileNameWithoutSuffix + ": 第4行" + checkedColIndex.ToString() + "列非数字;   导入取消。", false);
                    //timerRestoreTheLblResult.Start();
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
        /// 
        /// Attendance_Record_Detail
        /// </summary>
        /// <param name="attendanceMachineFlag"></param>
        /// <param name="year_and_month_str"></param>
        /// <returns></returns>
        /// 
    
        /// <summary>
        /// 删除本月的考勤记录 
        /// </summary>
        /// <param name="year_and_month_str"></param>
        private void delARDetailInfoByYearAndMonth(string year_and_month_str)
        {
            string sqlStr = string.Format(@"delete 
                                            from Attendance_record_Detail 
                                            where trunc(finger_print_date,'MM') = to_date('{0}','yyyy-MM')",
                                            year_and_month_str);
            OracleDaoHelper.executeSQL(sqlStr);
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
            return dt;
        }
        /// <summary>
        /// 获取汉字相同的记录的组数。
        /// </summary>
        /// <returns></returns>
        private int getAmountOfGroupOfSameName() {
            string sqlStr = string.Format(@"SELECT ""姓名""
                                            FROM
                                            (
                                                select distinct
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
                                            )   T
                                            group by ""姓名""");
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        private System.Data.DataTable getSamePinYinButName()
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
        /// 获取姓名拼音相同，但汉字书写不同的记录的组数。
        /// </summary>
        /// <returns></returns>
        private int getAmountOfGroupOfSamePinYinButName() {
            string sqlStr = string.Format(@"select 1
                                                    from 
                                                    (
                                                        select distinct 
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
                                                    ) T
                                                    group by gethzpy.GetHzFullPY(""姓名"")");
            return OracleDaoHelper.getDTBySql(sqlStr).Rows.Count;
        }
        private void saveCriticalARInfo(List<string> xlsFilePathList)
        {
            //先清除所有记录。
            AR_Temp.deleteTheARTemp();
            for(int i=0;i<=xlsFilePathList.Count-1;i++)
            {
                string excelPath = xlsFilePathList[i];
                //打开文档
                MyExcel myExcel = new MyExcel(excelPath);
                myExcel.open();
                Worksheet firstWS = myExcel.getFirstWorkSheetAfterOpen();
                //删除  时间后立即为空的行。
                AttendanceRHelper.clearSheet(firstWS);
                Usual_Excel_Helper uEHelper = new Usual_Excel_Helper(firstWS);
                string excelName = Usual_Excel_Helper.getExcelName(excelPath);
                //先获取第4行的最大行列数目。
                int rowMaxIndex = firstWS.UsedRange.Rows.Count;
                int pbMaximum = rowMaxIndex - 4;
                int pbValue = 0;
                
                //0: 表示 lblResult.Visible
                checkNameBGWorker.ReportProgress(0, "lblResult.Visible");
                checkNameBGWorker.ReportProgress(pbMaximum, "pb.Maximum");
                checkNameBGWorker.ReportProgress(pbValue, "pb.Value");
                //lblPrompt.Text = excelName + ": 基本信息采集中...";
                checkNameBGWorker.ReportProgress(0,string.Format(@"lblPrompt.Text={0}: 姓名采集中...",excelName));
                for (int rowIndex = 5; rowIndex <= rowMaxIndex; rowIndex++)
                {
                    //偶数行为 时间。
                    if (0 == rowIndex % 2)
                    {
                        checkNameBGWorker.ReportProgress(pbValue++, "pb.Value");
                        continue;
                    }
                    //姓名 存于第11列。
                    string name = uEHelper.getCellContentByRowAndColIndex(rowIndex, 11);
                    AR_Temp ar_Temp = new AR_Temp();
                    ar_Temp.Attendance_machine_flag = int.Parse(excelName.Substring(excelName.Length - 1, 1));
                    ar_Temp.Row_Index = rowIndex;
                    ar_Temp.Job_number = uEHelper.getCellContentByRowAndColIndex(rowIndex, 3);
                    ar_Temp.Name = name;
                    ar_Temp.saveRecord();
                    checkNameBGWorker.ReportProgress(pbValue++, "pb.Value");
                }
                myExcel.close();
            }
        }
        private void timerRestoreTheLblResult_Tick(object sender, EventArgs e)
        {
            timerRestoreTheLblResult.Enabled = false;
            lblResult.Text = "";
            lblResult.BackColor = this.BackColor;
            lblResult.Visible = false;
        }

        private void btnViewTheUncertaiRecordInExcel_Click(object sender, EventArgs e)
        {
            MyExcel uncertainRecordExcel = new MyExcel(_uncertainWSPath);
            uncertainRecordExcel.open(true);
        }


    }
}
